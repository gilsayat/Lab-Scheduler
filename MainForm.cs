using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;

namespace LabtechScheduler
{
    public class MainForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Core UI Controls
        private RadioButton rbNight;
        private RadioButton rbDay;
        private Label lblExcel;
        private Label lblDb;
        private Label lblStatus;
        private DataGridView gridPreview;
        private Button btnAutoFill;

        private string excelPath = "";
        private string dbPath = "Background Database HORIZON.txt";
        private List<string> dbEntries = new List<string>();

        // Parsing Constants
        private readonly string[] Gangways = { "ensenada", "mazatlan", "puerto vallarta", "la paz", "cabo", "cabo san lucas" };
        private readonly string[] Decks = { "deck", "waterworks", "water works", "slide", "slides", "skyride", "sky ride", "lido rail", "rail", "ropes", "water in the back" };
        private readonly string[] Costumes = { "elephant", "ellie", "duck", "ducky", "pirate" };
        private readonly string[] Dining = { "horizon", "vista", "steakhouse", "cucina", "pig & anchor", "pigandanchor", "alchemy", "coffee shop", "dining" };
        private readonly string[] NonShooting = { "gallery", "production", "edit till", "retail", "refill", "open gallery", "promote iris", "3 dreams booking", "organise and clean", "40+ sessions", "2 ds bookings" };

        private Dictionary<int, Tuple<string, string>> currentSchedule = new Dictionary<int, Tuple<string, string>>();

        public MainForm()
        {
            InitializeComponent();
            ApplyCustomSchedulerIcon();
            LoadDatabase();
        }

        private void InitializeComponent()
        {
            this.Text = "Labtech Scheduler";
            this.Size = new Size(650, 620);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // 1. Shift Selection Group
            GroupBox gbShift = new GroupBox
            {
                Text = "1. Event Shift Toggle",
                Location = new Point(15, 10),
                Size = new Size(605, 65)
            };

            rbNight = new RadioButton
            {
                Text = "🌙 Night Event (Portraits, Dining Rooms, 5 PM - 10 PM)",
                Location = new Point(20, 25),
                AutoSize = true,
                Checked = true
            };
            rbNight.CheckedChanged += (s, e) => { if (rbNight.Checked) ParseAndPreview(); };

            rbDay = new RadioButton
            {
                Text = "☀️ Day Event (Deck, Waterworks, Gangway, Costumes, Before 4 PM)",
                Location = new Point(20, 42),
                AutoSize = true
            };
            rbDay.CheckedChanged += (s, e) => { if (rbDay.Checked) ParseAndPreview(); };

            gbShift.Controls.Add(rbNight);
            gbShift.Controls.Add(rbDay);

            // 2. Schedule Excel Group
            GroupBox gbExcel = new GroupBox
            {
                Text = "2. Daily Excel Schedule (.xlsx)",
                Location = new Point(15, 80),
                Size = new Size(605, 60)
            };

            lblExcel = new Label
            {
                Text = "No Excel schedule file loaded...",
                Location = new Point(15, 25),
                Size = new Size(460, 20),
                ForeColor = Color.Gray
            };

            Button btnBrowseExcel = new Button
            {
                Text = "Browse Excel",
                Location = new Point(485, 20),
                Size = new Size(105, 28)
            };
            btnBrowseExcel.Click += BrowseExcel_Click;

            gbExcel.Controls.Add(lblExcel);
            gbExcel.Controls.Add(btnBrowseExcel);

            // 3. Database Manager Group
            GroupBox gbDb = new GroupBox
            {
                Text = "3. Background Database (.txt)",
                Location = new Point(15, 145),
                Size = new Size(605, 60)
            };

            lblDb = new Label
            {
                Text = "Database: Background Database HORIZON.txt (0 items)",
                Location = new Point(15, 25),
                Size = new Size(360, 20)
            };

            Button btnAddEntry = new Button
            {
                Text = "+ Add Entry",
                Location = new Point(380, 20),
                Size = new Size(100, 28)
            };
            btnAddEntry.Click += AddEntry_Click;

            Button btnChangeDb = new Button
            {
                Text = "Change File",
                Location = new Point(485, 20),
                Size = new Size(105, 28)
            };
            btnChangeDb.Click += ChangeDb_Click;

            gbDb.Controls.Add(lblDb);
            gbDb.Controls.Add(btnAddEntry);
            gbDb.Controls.Add(btnChangeDb);

            // 4. Preview Grid Group
            GroupBox gbPreview = new GroupBox
            {
                Text = "4. Schedule Preview (What will be auto-filled)",
                Location = new Point(15, 210),
                Size = new Size(605, 275)
            };

            gridPreview = new DataGridView
            {
                Location = new Point(15, 22),
                Size = new Size(575, 240),
                ReadOnly = true,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White
            };

            gridPreview.Columns.Add("Slot", "Slot");
            gridPreview.Columns.Add("Col2", "Role Code (Col 2)");
            gridPreview.Columns.Add("Col3", "Schedule / Station (Col 3)");

            gridPreview.Columns[0].Width = 70;
            gridPreview.Columns[1].Width = 140;
            gridPreview.Columns[2].Width = 360;

            gbPreview.Controls.Add(gridPreview);

            // 5. Actions / Status
            lblStatus = new Label
            {
                Text = "Ready. Select an Excel file and shift.",
                Location = new Point(15, 495),
                Size = new Size(605, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray
            };

            btnAutoFill = new Button
            {
                Text = "⚡ AUTO-FILL INTO TALLY BOARD ⚡",
                Location = new Point(15, 520),
                Size = new Size(605, 45),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnAutoFill.Click += AutoFill_Click;

            this.Controls.Add(gbShift);
            this.Controls.Add(gbExcel);
            this.Controls.Add(gbDb);
            this.Controls.Add(gbPreview);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnAutoFill);
        }

        private void ApplyCustomSchedulerIcon()
        {
            // Dynamically generate a clean "Lab Scheduler" icon in memory
            using (Bitmap bmp = new Bitmap(48, 48))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // Calendar Body
                g.FillRectangle(Brushes.White, 4, 8, 40, 36);
                g.FillRectangle(new SolidBrush(Color.FromArgb(0, 120, 215)), 4, 8, 40, 12);
                g.DrawRectangle(new Pen(Color.FromArgb(30, 30, 30), 2), 4, 8, 40, 36);

                // Clock Overlay on bottom right
                g.FillEllipse(new SolidBrush(Color.FromArgb(255, 185, 0)), 22, 22, 22, 22);
                g.DrawEllipse(new Pen(Color.FromArgb(40, 40, 40), 2), 22, 22, 22, 22);

                // Clock Hands
                g.DrawLine(new Pen(Color.FromArgb(30, 30, 30), 2), 33, 33, 33, 26);
                g.DrawLine(new Pen(Color.FromArgb(30, 30, 30), 2), 33, 33, 39, 33);

                IntPtr hIcon = bmp.GetHicon();
                this.Icon = Icon.FromHandle(hIcon);
            }
        }

        private void LoadDatabase()
        {
            dbEntries.Clear();
            if (File.Exists(dbPath))
            {
                string[] lines = File.ReadAllLines(dbPath);
                foreach (string line in lines)
                {
                    string trimmed = line.Trim();
                    if (!string.IsNullOrEmpty(trimmed) && trimmed.Contains(","))
                    {
                        dbEntries.Add(trimmed);
                    }
                }
            }
            lblDb.Text = $"Database: {Path.GetFileName(dbPath)} ({dbEntries.Count} items)";
        }

        private void BrowseExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    excelPath = ofd.FileName;
                    lblExcel.Text = Path.GetFileName(excelPath);
                    lblExcel.ForeColor = Color.Black;
                    ParseAndPreview();
                }
            }
        }

        private void ChangeDb_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Text Files|*.txt|All Files|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    dbPath = ofd.FileName;
                    LoadDatabase();
                    ParseAndPreview();
                    MessageBox.Show($"Loaded {dbEntries.Count} items from {Path.GetFileName(dbPath)}.", "Database Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void AddEntry_Click(object sender, EventArgs e)
        {
            string input = PromptInput("Add New Station / Background", "Enter format: Code,Name\n(e.g., 14036,Horizon)");
            if (!string.IsNullOrWhiteSpace(input) && input.Contains(","))
            {
                File.AppendAllText(dbPath, Environment.NewLine + input.Trim());
                LoadDatabase();
                ParseAndPreview();
                MessageBox.Show($"Added '{input.Trim()}' to {Path.GetFileName(dbPath)}!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string PromptInput(string title, string prompt)
        {
            Form promptForm = new Form
            {
                Width = 380,
                Height = 170,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = title,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false
            };
            Label textLabel = new Label { Left = 20, Top = 15, Text = prompt, AutoSize = true };
            TextBox textBox = new TextBox { Left = 20, Top = 55, Width = 320 };
            Button confirmation = new Button { Text = "Save", Left = 240, Width = 100, Top = 90, DialogResult = DialogResult.OK };
            promptForm.Controls.Add(textLabel);
            promptForm.Controls.Add(textBox);
            promptForm.Controls.Add(confirmation);
            promptForm.AcceptButton = confirmation;

            return promptForm.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }

        private string FuzzyMatchStation(string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return "";
            string clean = query.Trim().ToLower();

            // 1. Direct contains or code match
            foreach (string entry in dbEntries)
            {
                string[] parts = entry.Split(new[] { ',' }, 2);
                string code = parts[0].Trim().ToLower();
                string name = parts[1].Trim().ToLower();

                if (name == clean || entry.ToLower() == clean || code == clean)
                    return entry;

                if (clean.Length >= 4 && (name.Contains(clean) || clean.Contains(name)))
                    return entry;
            }

            // 2. Levenshtein / Similarity match (typo tolerance)
            double bestScore = 0.0;
            string bestMatch = "";

            foreach (string entry in dbEntries)
            {
                string[] parts = entry.Split(new[] { ',' }, 2);
                string name = parts[1].Trim().ToLower();

                double score = CalculateSimilarity(clean, name);
                if (score > bestScore)
                {
                    bestScore = score;
                    bestMatch = entry;
                }
            }

            if (bestScore >= 0.70)
                return bestMatch;

            return "";
        }

        private double CalculateSimilarity(string s, string t)
        {
            if (string.IsNullOrEmpty(s) || string.IsNullOrEmpty(t)) return 0.0;
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            for (int i = 0; i <= n; d[i, 0] = i++) { }
            for (int j = 0; j <= m; d[0, j] = j++) { }

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }

            int maxLen = Math.Max(n, m);
            return 1.0 - ((double)d[n, m] / maxLen);
        }

        private void ParseAndPreview()
        {
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath)) return;

            currentSchedule.Clear();
            gridPreview.Rows.Clear();

            bool isNight = rbNight.Checked;
            int startRow = isNight ? 7 : 0;
            int endRow = isNight ? 12 : 4;

            try
            {
                using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTable table = result.Tables[0];

                    // Read photographers from columns 2 to end
                    for (int c = 2; c < table.Columns.Count; c++)
                    {
                        string colHeader = table.Rows[0][c]?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(colHeader))
                            colHeader = table.Columns[c].ColumnName;

                        Match m = Regex.Match(colHeader, @"P(\d+)", RegexOptions.IgnoreCase);
                        if (!m.Success) continue;

                        int pNum = int.Parse(m.Groups[1].Value);

                        List<string> tasks = new List<string>();
                        for (int r = startRow; r <= endRow && r < table.Rows.Count; r++)
                        {
                            string val = table.Rows[r][c]?.ToString()?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(val) && !val.Equals("nan", StringComparison.OrdinalIgnoreCase))
                                tasks.Add(val);
                        }

                        Tuple<string, string> parsed = DetermineRoleAndStation(pNum, tasks, isNight);
                        currentSchedule[pNum] = parsed;
                    }
                }

                // Populate Grid for P01 to P30
                int activeCount = 0;
                for (int i = 1; i <= 30; i++)
                {
                    string role = "";
                    string station = "";

                    if (currentSchedule.ContainsKey(i))
                    {
                        role = currentSchedule[i].Item1;
                        station = currentSchedule[i].Item2;
                    }

                    gridPreview.Rows.Add($"P{i:D2}", role, station);
                    if (!string.IsNullOrEmpty(role)) activeCount++;
                }

                string shiftName = isNight ? "Night Event" : "Day Event";
                lblStatus.Text = $"{shiftName} Loaded: {activeCount} photographers assigned.";
                lblStatus.ForeColor = Color.DarkGreen;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Tuple<string, string> DetermineRoleAndStation(int pNum, List<string> tasks, bool isNight)
        {
            string combined = string.Join(" // ", tasks).ToLower();
            string role = "";
            string station = "";

            if (isNight)
            {
                // 1. Dining Rooms
                foreach (string dr in Dining)
                {
                    if (combined.Contains(dr))
                    {
                        role = $"DR{pNum}";
                        string match = FuzzyMatchStation(dr);
                        station = !string.IsNullOrEmpty(match) ? match : Capitalize(dr);
                        return Tuple.Create(role, station);
                    }
                }

                // 2. Costumes
                foreach (string c in Costumes)
                {
                    if (combined.Contains(c))
                    {
                        role = $"CS{pNum}";
                        string match = FuzzyMatchStation(c);
                        station = !string.IsNullOrEmpty(match) ? match : Capitalize(c);
                        return Tuple.Create(role, station);
                    }
                }

                // 3. Studio Backdrops / Locations
                bool isNonShooting = false;
                foreach (string ns in new[] { "edit till", "gallery 4 pm", "promote iris", "refill retails" })
                {
                    if (combined.Contains(ns) && !combined.Contains("backdrop") && !combined.Contains("sequin") && !combined.Contains("mystique") && !combined.Contains("gate"))
                    {
                        isNonShooting = true;
                        break;
                    }
                }

                if (!isNonShooting)
                {
                    for (int i = tasks.Count - 1; i >= 0; i--)
                    {
                        string candidate = tasks[i].Trim();
                        if (candidate.Length > 2 && !ContainsAny(candidate.ToLower(), new[] { "team set up", "booking", "sessions", "working hours", "rest hours", "40+ sessions" }))
                        {
                            string matched = FuzzyMatchStation(candidate);
                            if (!string.IsNullOrEmpty(matched))
                            {
                                return Tuple.Create($"P{pNum}", matched);
                            }
                            else if (string.IsNullOrEmpty(station) && !ContainsAny(candidate.ToLower(), new[] { "bring", "flash", "poses", "sitting", "standing", "images", "call", "replace" }))
                            {
                                station = candidate;
                                role = $"P{pNum}";
                            }
                        }
                    }
                }
            }
            else
            {
                // DAY EVENT (Before 4 PM)
                // 1. Costumes
                foreach (string c in Costumes)
                {
                    if (combined.Contains(c))
                    {
                        role = $"CS{pNum}";
                        string match = FuzzyMatchStation(c);
                        station = !string.IsNullOrEmpty(match) ? match : Capitalize(c);
                        return Tuple.Create(role, station);
                    }
                }

                // 2. Gangways
                foreach (string gw in Gangways)
                {
                    if (combined.Contains(gw))
                    {
                        role = $"GW{pNum}";
                        string match = FuzzyMatchStation(gw);
                        station = !string.IsNullOrEmpty(match) ? match : Capitalize(gw);
                        return Tuple.Create(role, station);
                    }
                }

                // 3. Decks / Outdoors
                foreach (string dk in Decks)
                {
                    if (combined.Contains(dk))
                    {
                        role = $"DK{pNum}";
                        string match = FuzzyMatchStation(dk);
                        station = !string.IsNullOrEmpty(match) ? match : Capitalize(dk);
                        return Tuple.Create(role, station);
                    }
                }

                // 4. Portrait/Backdrop in Day
                foreach (string t in tasks)
                {
                    if (!ContainsAny(t.ToLower(), NonShooting) && !ContainsAny(t.ToLower(), new[] { "img", "target", "organise", "couples & singles" }))
                    {
                        string match = FuzzyMatchStation(t);
                        if (!string.IsNullOrEmpty(match))
                        {
                            return Tuple.Create($"P{pNum}", match);
                        }
                    }
                }
            }

            return Tuple.Create(role, station);
        }

        private bool ContainsAny(string str, string[] list)
        {
            foreach (string item in list)
            {
                if (str.Contains(item)) return true;
            }
            return false;
        }

        private string Capitalize(string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            return char.ToUpper(str[0]) + str.Substring(1);
        }

        private async void AutoFill_Click(object sender, EventArgs e)
        {
            if (currentSchedule.Count == 0)
            {
                MessageBox.Show("Please select an Excel schedule file first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string shift = rbNight.Checked ? "NIGHT" : "DAY";

            var confirm = MessageBox.Show(
                $"Auto-Fill [{shift}] Schedule into Tally Board:\n\n" +
                "1. Click OK.\n" +
                "2. Switch to 'SCHEDULER TALLY BOARD' within 4 seconds.\n" +
                "3. Click inside the Column 2 text box for Row P01.\n" +
                "4. Hands off the keyboard & mouse!",
                "Start Auto-Fill",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information
            );

            if (confirm != DialogResult.OK) return;

            lblStatus.Text = "Starting in 4 seconds... Click inside P01 Column 2 in Tally Board!";
            lblStatus.ForeColor = Color.Red;
            btnAutoFill.Enabled = false;

            await Task.Delay(4000);

            // Attempt to bring TallyBoard to the front automatically
            IntPtr tallyHwnd = FindWindow(null, "SCHEDULER TALLY BOARD V5.93");
            if (tallyHwnd == IntPtr.Zero)
            {
                // Try partial search or just foreground window
                foreach (var proc in System.Diagnostics.Process.GetProcesses())
                {
                    if (proc.MainWindowTitle.Contains("SCHEDULER TALLY BOARD"))
                    {
                        SetForegroundWindow(proc.MainWindowHandle);
                        break;
                    }
                }
            }
            else
            {
                SetForegroundWindow(tallyHwnd);
            }

            await Task.Delay(500);

            lblStatus.Text = "Typing schedule data into Tally Board...";
            lblStatus.ForeColor = Color.Blue;

            for (int i = 1; i <= 30; i++)
            {
                string role = "";
                string station = "";

                if (currentSchedule.ContainsKey(i))
                {
                    role = currentSchedule[i].Item1;
                    station = currentSchedule[i].Item2;
                }

                // 1. Column 2 (Role)
                if (!string.IsNullOrEmpty(role))
                {
                    Clipboard.SetText(role);
                    SendKeys.SendWait("^v");
                }
                await Task.Delay(40);
                SendKeys.SendWait("{TAB}");

                // 2. Column 3 (Station)
                if (!string.IsNullOrEmpty(station))
                {
                    Clipboard.SetText(station);
                    SendKeys.SendWait("^v");
                }
                await Task.Delay(40);
                SendKeys.SendWait("{TAB}");

                // 3. Move past Column 4 to next row
                SendKeys.SendWait("{TAB}");
                await Task.Delay(30);
            }

            btnAutoFill.Enabled = true;
            lblStatus.Text = $"{shift} Schedule filled successfully!";
            lblStatus.ForeColor = Color.DarkGreen;
            MessageBox.Show($"{shift} Schedule filled successfully into Tally Board!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
