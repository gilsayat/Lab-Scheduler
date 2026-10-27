using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
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

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private RadioButton rbNight;
        private RadioButton rbDay;
        private ComboBox cbDaySelect;
        private Label lblFile;
        private Label lblDb;
        private Label lblStatus;
        private DataGridView gridPreview;
        private Button btnAutoFill;

        private string loadedFilePath = "";
        private bool isImageFile = false;
        private string dbPath = "Background Database HORIZON.txt";
        private string rosterPath = "PhotographerRoster.txt";
        private List<string> dbEntries = new List<string>();
        private Dictionary<string, int> customNameMappings = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        private readonly string[] PhysicalLocations = {
            "steakhouse", "coffee shop", "pig & anchor", "pig and anchor", "pigandanchor",
            "dreamscape", "cherry on top", "liquor shop", "liquour shop",
            "gallery 1", "gallery 2", "open gallery", "gallery",
            "lobby dreams", "lobby gs1", "lobby gs2", "lobby gs",
            "deck 3", "deck 4", "deck 5", "deck staircase", "staircase starts", "staircase"
        };

        private readonly string[] IgnoreKeywords = {
            "off", "evening off", "day off", "vacation", "sick",
            "props", "props bag", "bag 1", "bag 2", "barriers", "bench", "tables", "stand/sit",
            "working hours", "rest hours", "sessions", "target", "booking", "challenge",
            "team set up", "tear down", "emergency tear down", "port manning", "clean & flyers",
            "edit till", "promote iris", "refill", "retail", "production", "organise"
        };

        private readonly string[] Costumes = { "elephant", "ellie", "duck", "ducky", "pirate", "bunny", "santa", "sailor" };
        private readonly string[] Gangways = { "ensenada", "mazatlan", "puerto vallarta", "la paz", "cabo", "cabo san lucas", "cozumel", "amber cove" };
        private readonly string[] Decks = { "deck", "waterworks", "water works", "slide", "slides", "skyride", "sky ride", "lido rail", "rail", "ropes" };

        private Dictionary<int, Tuple<string, string>> currentSchedule = new Dictionary<int, Tuple<string, string>>();

        public MainForm()
        {
            InitializeComponent();
            ApplyLtsIcon();
            TryFindDatabaseFile();
            LoadDatabase();
            LoadRosterSettings();
            EnableDragAndDrop();
        }

        private void TryFindDatabaseFile()
        {
            if (File.Exists("PO Background Database.txt"))
                dbPath = "PO Background Database.txt";
            else if (File.Exists("Background Database HORIZON.txt"))
                dbPath = "Background Database HORIZON.txt";
        }

        private void LoadRosterSettings()
        {
            customNameMappings.Clear();
            if (!File.Exists(rosterPath))
            {
                // Default setup: Gil -> 14, Sanyukta -> 15
                customNameMappings["Gil"] = 14;
                customNameMappings["Sanyukta"] = 15;
                SaveRosterSettings();
                return;
            }

            string[] lines = File.ReadAllLines(rosterPath);
            foreach (string line in lines)
            {
                string[] parts = line.Split('=');
                if (parts.Length == 2 && int.TryParse(parts[0].Trim(), out int slot))
                {
                    string name = parts[1].Trim();
                    if (!string.IsNullOrEmpty(name))
                    {
                        customNameMappings[name] = slot;
                    }
                }
            }
        }

        private void SaveRosterSettings()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var kvp in customNameMappings)
            {
                sb.AppendLine($"{kvp.Value}={kvp.Key}");
            }
            File.WriteAllText(rosterPath, sb.ToString());
        }

        private void InitializeComponent()
        {
            this.Text = "Labtech Scheduler v1.6";
            this.Size = new Size(660, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // 1. Shift & Day
            GroupBox gbShift = new GroupBox
            {
                Text = "1. Shift & Cruise Day Selection",
                Location = new Point(15, 10),
                Size = new Size(615, 75)
            };

            rbNight = new RadioButton
            {
                Text = "🌙 Night Event (Portraits, Dining)",
                Location = new Point(15, 22),
                AutoSize = true,
                Checked = true
            };
            rbNight.CheckedChanged += (s, e) => { if (rbNight.Checked) ParseCurrentFile(); };

            rbDay = new RadioButton
            {
                Text = "☀️ Day Event (Deck, Gangway, Costumes)",
                Location = new Point(15, 45),
                AutoSize = true
            };
            rbDay.CheckedChanged += (s, e) => { if (rbDay.Checked) ParseCurrentFile(); };

            Label lblDayPick = new Label
            {
                Text = "Cruise Day (Sequential Station Match):",
                Location = new Point(330, 22),
                AutoSize = true
            };

            cbDaySelect = new ComboBox
            {
                Location = new Point(330, 42),
                Size = new Size(160, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cbDaySelect.Items.AddRange(new object[] { "Day 1", "Day 2", "Day 3", "Day 4", "Day 5", "Day 6", "Day 7", "Day 8" });
            cbDaySelect.SelectedIndex = 0;
            cbDaySelect.SelectedIndexChanged += (s, e) => ParseCurrentFile();

            Button btnRosterSettings = new Button
            {
                Text = "⚙️ Roster",
                Location = new Point(500, 40),
                Size = new Size(95, 28)
            };
            btnRosterSettings.Click += (s, e) => OpenRosterSettingsDialog();

            gbShift.Controls.Add(rbNight);
            gbShift.Controls.Add(rbDay);
            gbShift.Controls.Add(lblDayPick);
            gbShift.Controls.Add(cbDaySelect);
            gbShift.Controls.Add(btnRosterSettings);

            // 2. Schedule File
            GroupBox gbFile = new GroupBox
            {
                Text = "2. Schedule Source (.xlsx, .jpg, .png) [Drag & Drop Supported]",
                Location = new Point(15, 90),
                Size = new Size(615, 60)
            };

            lblFile = new Label
            {
                Text = "Drag & drop an Excel schedule or Photo here...",
                Location = new Point(15, 25),
                Size = new Size(470, 20),
                ForeColor = Color.Gray
            };

            Button btnBrowse = new Button
            {
                Text = "Browse File",
                Location = new Point(495, 20),
                Size = new Size(105, 28)
            };
            btnBrowse.Click += BrowseFile_Click;

            gbFile.Controls.Add(lblFile);
            gbFile.Controls.Add(btnBrowse);

            // 3. Database
            GroupBox gbDb = new GroupBox
            {
                Text = "3. Background Database (.txt)",
                Location = new Point(15, 155),
                Size = new Size(615, 60)
            };

            lblDb = new Label
            {
                Text = $"Database: {Path.GetFileName(dbPath)} (0 items)",
                Location = new Point(15, 25),
                Size = new Size(370, 20)
            };

            Button btnAddEntry = new Button
            {
                Text = "+ Add Entry",
                Location = new Point(390, 20),
                Size = new Size(100, 28)
            };
            btnAddEntry.Click += AddEntry_Click;

            Button btnChangeDb = new Button
            {
                Text = "Change File",
                Location = new Point(495, 20),
                Size = new Size(105, 28)
            };
            btnChangeDb.Click += ChangeDb_Click;

            gbDb.Controls.Add(lblDb);
            gbDb.Controls.Add(btnAddEntry);
            gbDb.Controls.Add(btnChangeDb);

            // 4. Preview Table
            GroupBox gbPreview = new GroupBox
            {
                Text = "4. Schedule Preview (Editable - Double-click any cell to adjust)",
                Location = new Point(15, 220),
                Size = new Size(615, 285)
            };

            gridPreview = new DataGridView
            {
                Location = new Point(15, 22),
                Size = new Size(585, 250),
                ReadOnly = false,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                BackgroundColor = Color.White
            };

            gridPreview.Columns.Add("Slot", "Slot");
            gridPreview.Columns.Add("Col2", "Role Code (Col 2)");
            gridPreview.Columns.Add("Col3", "Schedule / Station (Col 3)");

            gridPreview.Columns[0].ReadOnly = true;
            gridPreview.Columns[0].Width = 70;
            gridPreview.Columns[1].Width = 140;
            gridPreview.Columns[2].Width = 370;

            gbPreview.Controls.Add(gridPreview);

            // 5. Actions / Status
            lblStatus = new Label
            {
                Text = "Ready. Drag an Excel schedule file or photo here to begin.",
                Location = new Point(15, 510),
                Size = new Size(615, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray
            };

            btnAutoFill = new Button
            {
                Text = "⚡ AUTO-FILL INTO TALLY BOARD (CLICK-BASED) ⚡",
                Location = new Point(15, 535),
                Size = new Size(615, 45),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnAutoFill.Click += AutoFill_Click;

            this.Controls.Add(gbShift);
            this.Controls.Add(gbFile);
            this.Controls.Add(gbDb);
            this.Controls.Add(gbPreview);
            this.Controls.Add(lblStatus);
            this.Controls.Add(btnAutoFill);
        }

        private void OpenRosterSettingsDialog()
        {
            Form settingsForm = new Form
            {
                Text = "Photographer Roster Settings",
                Size = new Size(400, 520),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            Panel scrollPanel = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(365, 410),
                AutoScroll = true
            };

            Dictionary<int, TextBox> boxes = new Dictionary<int, TextBox>();
            int yPos = 10;

            // Invert lookup
            Dictionary<int, string> slotToName = new Dictionary<int, string>();
            foreach (var kvp in customNameMappings)
                slotToName[kvp.Value] = kvp.Key;

            for (int i = 1; i <= 30; i++)
            {
                Label lbl = new Label { Text = $"Slot P{i:D2}:", Location = new Point(10, yPos + 3), Width = 80 };
                TextBox txt = new TextBox { Location = new Point(100, yPos), Width = 230 };

                if (slotToName.ContainsKey(i))
                    txt.Text = slotToName[i];

                boxes[i] = txt;
                scrollPanel.Controls.Add(lbl);
                scrollPanel.Controls.Add(txt);
                yPos += 30;
            }

            Button btnSave = new Button
            {
                Text = "Save Roster",
                Location = new Point(130, 430),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnSave.Click += (s, e) =>
            {
                customNameMappings.Clear();
                foreach (var kvp in boxes)
                {
                    string name = kvp.Value.Text.Trim();
                    if (!string.IsNullOrEmpty(name))
                    {
                        customNameMappings[name] = kvp.Key;
                    }
                }
                SaveRosterSettings();
                settingsForm.Close();
                ParseCurrentFile();
                MessageBox.Show("Photographer roster settings saved successfully!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            settingsForm.Controls.Add(scrollPanel);
            settingsForm.Controls.Add(btnSave);
            settingsForm.ShowDialog();
        }

        private void ApplyLtsIcon()
        {
            using (Bitmap bmp = new Bitmap(48, 48))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                using (GraphicsPath path = new GraphicsPath())
                {
                    int radius = 10;
                    path.AddArc(2, 2, radius, radius, 180, 90);
                    path.AddArc(46 - radius, 2, radius, radius, 270, 90);
                    path.AddArc(46 - radius, 46 - radius, radius, radius, 0, 90);
                    path.AddArc(2, 46 - radius, radius, radius, 90, 90);
                    path.CloseFigure();

                    using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(24, 30, 44)))
                    {
                        g.FillPath(bgBrush, path);
                    }
                    using (Pen borderPen = new Pen(Color.FromArgb(0, 150, 255), 2))
                    {
                        g.DrawPath(borderPen, path);
                    }
                }

                using (Font font = new Font("Segoe UI", 13, FontStyle.Bold, GraphicsUnit.Pixel))
                using (StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                {
                    g.DrawString("LTS", font, Brushes.White, new RectangleF(0, 0, 48, 48), sf);
                }

                IntPtr hIcon = bmp.GetHicon();
                this.Icon = Icon.FromHandle(hIcon);
            }
        }

        private void EnableDragAndDrop()
        {
            this.AllowDrop = true;
            this.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };

            this.DragDrop += (s, e) =>
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    ProcessDroppedFile(files[0]);
                }
            };
        }

        private void ProcessDroppedFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();

            if (ext == ".xlsx" || ext == ".xls")
            {
                isImageFile = false;
                loadedFilePath = filePath;
                lblFile.Text = Path.GetFileName(loadedFilePath);
                lblFile.ForeColor = Color.Black;
                ParseCurrentFile();
            }
            else if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp")
            {
                isImageFile = true;
                loadedFilePath = filePath;
                lblFile.Text = "[Photo] " + Path.GetFileName(loadedFilePath);
                lblFile.ForeColor = Color.Black;
                ParseCurrentFile();
            }
            else if (ext == ".txt")
            {
                dbPath = filePath;
                LoadDatabase();
                ParseCurrentFile();
                MessageBox.Show($"Loaded database: {Path.GetFileName(dbPath)}", "Database Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BrowseFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Schedule Files|*.xlsx;*.xls;*.jpg;*.jpeg;*.png;*.bmp|Excel Files|*.xlsx;*.xls|Images|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    ProcessDroppedFile(ofd.FileName);
                }
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

        private void ChangeDb_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Text Files|*.txt|All Files|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    dbPath = ofd.FileName;
                    LoadDatabase();
                    ParseCurrentFile();
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
                ParseCurrentFile();
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

        private string GetDiningStationByDay(string diningName, int dayIndex)
        {
            string clean = diningName.Trim().ToLower();
            List<string> matches = new List<string>();

            foreach (string entry in dbEntries)
            {
                string[] parts = entry.Split(new[] { ',' }, 2);
                if (parts.Length == 2 && parts[1].Trim().ToLower() == clean)
                {
                    matches.Add(entry);
                }
            }

            if (matches.Count > 0)
            {
                int idx = (dayIndex - 1) % matches.Count;
                return matches[idx];
            }

            return "";
        }

        private string MatchBackgroundFromDatabase(string query, int dayIndex)
        {
            if (string.IsNullOrWhiteSpace(query)) return "";
            string clean = Regex.Replace(query.Trim().ToLower(), @"\d{1,2}(:\d{2})?\s*(am|pm)?\s*(to|-)?\s*\d{1,2}(:\d{2})?\s*(am|pm)?", "").Trim();
            clean = Regex.Replace(clean, @"[^a-z0-9]", "");

            if (clean.Length < 3) return "";

            foreach (string loc in PhysicalLocations)
            {
                string cleanLoc = Regex.Replace(loc, @"[^a-z0-9]", "");
                if (clean == cleanLoc && clean != "alchemy")
                {
                    return "";
                }
            }

            List<string> exactMatches = new List<string>();

            foreach (string entry in dbEntries)
            {
                string[] parts = entry.Split(new[] { ',' }, 2);
                if (parts.Length < 2) continue;

                string nameClean = Regex.Replace(parts[1].Trim().ToLower(), @"[^a-z0-9]", "");

                if (nameClean == clean || clean.Contains(nameClean) || nameClean.Contains(clean))
                {
                    exactMatches.Add(entry);
                }
            }

            if (exactMatches.Count > 0)
            {
                int idx = (dayIndex - 1) % exactMatches.Count;
                return exactMatches[idx];
            }

            double bestScore = 0.0;
            string bestMatch = "";

            foreach (string entry in dbEntries)
            {
                string[] parts = entry.Split(new[] { ',' }, 2);
                if (parts.Length < 2) continue;

                string nameClean = Regex.Replace(parts[1].Trim().ToLower(), @"[^a-z0-9]", "");
                double score = CalculateSimilarity(clean, nameClean);
                if (score > bestScore)
                {
                    bestScore = score;
                    bestMatch = entry;
                }
            }

            if (bestScore >= 0.75)
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

        private void ParseCurrentFile()
        {
            if (string.IsNullOrEmpty(loadedFilePath) || !File.Exists(loadedFilePath)) return;

            if (isImageFile)
                ParseImageFile();
            else
                ParseExcelFile();
        }

        private void ParseExcelFile()
        {
            currentSchedule.Clear();
            gridPreview.Rows.Clear();

            bool isNight = rbNight.Checked;
            int dayIndex = cbDaySelect.SelectedIndex + 1;

            try
            {
                using (var stream = File.Open(loadedFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTable table = result.Tables[0];

                    int offsetRow = 0;
                    int startRow = offsetRow + (isNight ? 7 : 0);
                    int endRow = offsetRow + (isNight ? 12 : 4);

                    for (int c = 2; c < table.Columns.Count; c++)
                    {
                        string colHeader = table.Rows[offsetRow][c]?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(colHeader))
                            colHeader = table.Columns[c].ColumnName;

                        Match m = Regex.Match(colHeader, @"P(\d+)", RegexOptions.IgnoreCase);
                        if (!m.Success) continue;

                        int defaultSlot = int.Parse(m.Groups[1].Value);

                        List<string> tasks = new List<string>();
                        for (int r = startRow; r <= endRow && r < table.Rows.Count; r++)
                        {
                            string val = table.Rows[r][c]?.ToString()?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(val) && !val.Equals("nan", StringComparison.OrdinalIgnoreCase))
                                tasks.Add(val);
                        }

                        // Check name overrides
                        int finalSlot = defaultSlot;
                        string allText = colHeader + " " + string.Join(" ", tasks);

                        foreach (var kvp in customNameMappings)
                        {
                            if (Regex.IsMatch(allText, @"\b" + Regex.Escape(kvp.Key) + @"\b", RegexOptions.IgnoreCase))
                            {
                                finalSlot = kvp.Value;
                                break;
                            }
                        }

                        Tuple<string, string> parsed = DetermineRoleAndStation(finalSlot, tasks, isNight, dayIndex);
                        currentSchedule[finalSlot] = parsed;
                    }
                }

                DisplayScheduleInGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ParseImageFile()
        {
            lblStatus.Text = "Scanning photo with Windows OCR...";
            lblStatus.ForeColor = Color.Blue;
            this.Update();

            string fullText = RunWindowsNativeOcr(loadedFilePath);
            if (string.IsNullOrWhiteSpace(fullText))
            {
                MessageBox.Show("Could not extract text from the photo. Make sure the schedule is clear and visible.", "OCR Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lblStatus.Text = "OCR scan yielded no text. Try a clearer photo.";
                lblStatus.ForeColor = Color.Red;
                return;
            }

            currentSchedule.Clear();
            gridPreview.Rows.Clear();

            bool isNight = rbNight.Checked;
            int dayIndex = cbDaySelect.SelectedIndex + 1;
            string[] lines = fullText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                Match m = Regex.Match(line, @"\bP(\d+)\b", RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    int defaultSlot = int.Parse(m.Groups[1].Value);
                    List<string> contextLines = new List<string> { line };
                    for (int j = 1; j <= 3 && (i + j) < lines.Length; j++)
                    {
                        contextLines.Add(lines[i + j]);
                    }

                    int finalSlot = defaultSlot;
                    string allText = string.Join(" ", contextLines);

                    foreach (var kvp in customNameMappings)
                    {
                        if (Regex.IsMatch(allText, @"\b" + Regex.Escape(kvp.Key) + @"\b", RegexOptions.IgnoreCase))
                        {
                            finalSlot = kvp.Value;
                            break;
                        }
                    }

                    Tuple<string, string> parsed = DetermineRoleAndStation(finalSlot, contextLines, isNight, dayIndex);
                    if (!string.IsNullOrEmpty(parsed.Item1) || !string.IsNullOrEmpty(parsed.Item2))
                    {
                        currentSchedule[finalSlot] = parsed;
                    }
                }
            }

            DisplayScheduleInGrid();
        }

        private string RunWindowsNativeOcr(string imagePath)
        {
            try
            {
                string script = @"
$file = [System.IO.Path]::GetFullPath('" + imagePath.Replace("'", "''") + @"')
if (-not (Test-Path $file)) { exit }
[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime] | Out-Null
$asyncOp = [Windows.Storage.StorageFile]::GetFileFromPathAsync($file)
$storageFile = $asyncOp.GetAwaiter().GetResult()
$streamOp = $storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)
$stream = $streamOp.GetAwaiter().GetResult()
$decoderOp = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
$decoder = $decoderOp.GetAwaiter().GetResult()
$bitmapOp = $decoder.GetSoftwareBitmapAsync()
$bitmap = $bitmapOp.GetAwaiter().GetResult()
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($engine -eq $null) {
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('en-US'))
}
$ocrOp = $engine.RecognizeAsync($bitmap)
$result = $ocrOp.GetAwaiter().GetResult()
foreach ($line in $result.Lines) {
    Write-Output $line.Text
}
";
                string tempPs1 = Path.Combine(Path.GetTempPath(), "tally_ocr_" + Guid.NewGuid().ToString("N") + ".ps1");
                File.WriteAllText(tempPs1, script, Encoding.UTF8);

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + tempPs1 + "\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process proc = Process.Start(psi))
                {
                    string output = proc.StandardOutput.ReadToEnd();
                    proc.WaitForExit(10000);

                    try { File.Delete(tempPs1); } catch { }
                    return output;
                }
            }
            catch
            {
                return "";
            }
        }

        private void DisplayScheduleInGrid()
        {
            gridPreview.Rows.Clear();
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

            string shiftName = rbNight.Checked ? "Night Event" : "Day Event";
            lblStatus.Text = $"{shiftName} Loaded: {activeCount} photographers assigned. Double-click any cell to edit.";
            lblStatus.ForeColor = Color.DarkGreen;
        }

        private Tuple<string, string> DetermineRoleAndStation(int pNum, List<string> tasks, bool isNight, int dayIndex)
        {
            string combined = string.Join(" // ", tasks).ToLower();

            // 1. Off-duty
            foreach (string word in new[] { "evening off", "day off", "off oct", "off", "vacation", "sick" })
            {
                if (Regex.IsMatch(combined, @"\b" + word + @"\b"))
                {
                    return Tuple.Create("", "");
                }
            }

            // 2. Dining
            if (combined.Contains("horizon"))
            {
                string matched = GetDiningStationByDay("horizon", dayIndex);
                return Tuple.Create($"DR{pNum}", matched);
            }
            if (combined.Contains("vista"))
            {
                string matched = GetDiningStationByDay("vista", dayIndex);
                return Tuple.Create($"DR{pNum}", matched);
            }

            // 3. Costumes
            foreach (string c in Costumes)
            {
                if (combined.Contains(c))
                {
                    string matched = MatchBackgroundFromDatabase(c, dayIndex);
                    return Tuple.Create($"CS{pNum}", matched);
                }
            }

            // 4. DAY EVENT (Decks / Gangways)
            if (!isNight)
            {
                foreach (string gw in Gangways)
                {
                    if (combined.Contains(gw))
                    {
                        string matched = MatchBackgroundFromDatabase(gw, dayIndex);
                        return Tuple.Create($"GW{pNum}", matched);
                    }
                }

                foreach (string dk in Decks)
                {
                    if (combined.Contains(dk))
                    {
                        string matched = MatchBackgroundFromDatabase(dk, dayIndex);
                        return Tuple.Create($"DK{pNum}", matched);
                    }
                }
            }

            // 5. Studio Backdrops
            for (int i = tasks.Count - 1; i >= 0; i--)
            {
                string candidate = tasks[i].Trim();
                string candLower = candidate.ToLower();

                bool isPureLocation = false;
                foreach (string loc in PhysicalLocations)
                {
                    if (candLower == loc || (candLower.Contains(loc) && !candLower.Contains("alchemy") && !candLower.Contains("sequin") && !candLower.Contains("mystique")))
                    {
                        isPureLocation = true;
                        break;
                    }
                }
                if (isPureLocation) continue;

                if (ContainsAny(candLower, IgnoreKeywords))
                    continue;

                string matched = MatchBackgroundFromDatabase(candidate, dayIndex);
                if (!string.IsNullOrEmpty(matched))
                {
                    return Tuple.Create($"P{pNum}", matched);
                }
            }

            return Tuple.Create("", "");
        }

        private bool ContainsAny(string str, string[] list)
        {
            foreach (string item in list)
            {
                if (str.Contains(item)) return true;
            }
            return false;
        }

        private async void AutoFill_Click(object sender, EventArgs e)
        {
            for (int r = 0; r < gridPreview.Rows.Count; r++)
            {
                int pNum = r + 1;
                string role = gridPreview.Rows[r].Cells[1].Value?.ToString() ?? "";
                string station = gridPreview.Rows[r].Cells[2].Value?.ToString() ?? "";
                currentSchedule[pNum] = Tuple.Create(role, station);
            }

            string shift = rbNight.Checked ? "NIGHT" : "DAY";

            var confirm = MessageBox.Show(
                $"Auto-Fill [{shift}] Schedule into Tally Board:\n\n" +
                "1. Click OK.\n" +
                "2. Click on the Tally Board window during the 3-second countdown to ensure it has focus.\n" +
                "3. Hands off mouse & keyboard!",
                "Start Auto-Fill",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information
            );

            if (confirm != DialogResult.OK) return;

            lblStatus.Text = "Targeting TallyBoard... (3 seconds)";
            lblStatus.ForeColor = Color.Red;
            btnAutoFill.Enabled = false;

            await Task.Delay(3000);

            IntPtr tallyHwnd = IntPtr.Zero;

            // Direct process name search for TallyBoard.exe
            Process[] procs = Process.GetProcessesByName("TallyBoard");
            if (procs.Length == 0)
                procs = Process.GetProcessesByName("tallyboard");

            if (procs.Length > 0)
            {
                tallyHwnd = procs[0].MainWindowHandle;
                SetForegroundWindow(tallyHwnd);
            }
            else
            {
                // Fallback: Check whatever window is currently foreground
                tallyHwnd = GetForegroundWindow();
            }

            if (tallyHwnd == IntPtr.Zero)
            {
                MessageBox.Show("Could not attach to TallyBoard. Please ensure TallyBoard is running.", "Process Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                btnAutoFill.Enabled = true;
                return;
            }

            await Task.Delay(500);

            RECT rect;
            GetWindowRect(tallyHwnd, out rect);
            int windowX = rect.Left;
            int windowY = rect.Top;

            int startRowY = windowY + 115;
            int rowHeight = 24;
            int col2X = windowX + 85;
            int col3X = windowX + 220;

            lblStatus.Text = "Click-filling schedule into Tally Board...";
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

                int currentY = startRowY + ((i - 1) * rowHeight);

                // Column 2 (Role)
                if (!string.IsNullOrEmpty(role))
                {
                    SetCursorPos(col2X, currentY);
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    await Task.Delay(25);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    await Task.Delay(40);

                    Clipboard.SetText(role);
                    SendKeys.SendWait("^v");
                    await Task.Delay(50);
                }

                // Column 3 (Station)
                if (!string.IsNullOrEmpty(station))
                {
                    SetCursorPos(col3X, currentY);
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    await Task.Delay(25);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    await Task.Delay(40);

                    Clipboard.SetText(station);
                    SendKeys.SendWait("^v");
                    await Task.Delay(60);
                }
            }

            btnAutoFill.Enabled = true;
            lblStatus.Text = $"{shift} Schedule filled successfully!";
            lblStatus.ForeColor = Color.DarkGreen;
            MessageBox.Show($"{shift} Schedule filled successfully into Tally Board!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
