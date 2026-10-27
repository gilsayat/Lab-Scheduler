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

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Core Controls
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
        private List<string> dbEntries = new List<string>();

        // Parsing Categories
        private readonly string[] Gangways = { "ensenada", "mazatlan", "puerto vallarta", "la paz", "cabo", "cabo san lucas" };
        private readonly string[] Decks = { "deck", "waterworks", "water works", "slide", "slides", "skyride", "sky ride", "lido rail", "rail", "ropes", "water in the back" };
        private readonly string[] Costumes = { "elephant", "ellie", "duck", "ducky", "pirate" };
        private readonly string[] Dining = { "horizon", "vista", "steakhouse", "cucina", "pig & anchor", "pigandanchor", "alchemy", "coffee shop", "dining" };
        private readonly string[] NonShooting = { "gallery", "production", "edit till", "retail", "refill", "open gallery", "promote iris", "3 dreams booking", "organise and clean", "40+ sessions", "2 ds bookings" };

        private Dictionary<int, Tuple<string, string>> currentSchedule = new Dictionary<int, Tuple<string, string>>();

        public MainForm()
        {
            InitializeComponent();
            ApplyLtsIcon();
            LoadDatabase();
            EnableDragAndDrop();
        }

        private void InitializeComponent()
        {
            this.Text = "Labtech Scheduler v1.3";
            this.Size = new Size(660, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // 1. Shift & Day Selection Group
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
                Text = "Target Day / Tab:",
                Location = new Point(350, 22),
                AutoSize = true
            };

            cbDaySelect = new ComboBox
            {
                Location = new Point(350, 42),
                Size = new Size(245, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cbDaySelect.Items.AddRange(new object[] { "Auto / Default", "Day 1", "Day 2", "Day 3", "Day 4", "Day 5", "Day 6", "Day 7", "Day 8" });
            cbDaySelect.SelectedIndex = 0;
            cbDaySelect.SelectedIndexChanged += (s, e) => ParseCurrentFile();

            gbShift.Controls.Add(rbNight);
            gbShift.Controls.Add(rbDay);
            gbShift.Controls.Add(lblDayPick);
            gbShift.Controls.Add(cbDaySelect);

            // 2. Schedule File Input Group (Drag and Drop Supported)
            GroupBox gbFile = new GroupBox
            {
                Text = "2. Schedule Source (.xlsx, .jpg, .png) [Drag & Drop Supported]",
                Location = new Point(15, 90),
                Size = new Size(615, 60)
            };

            lblFile = new Label
            {
                Text = "Drag & drop an Excel file or Photo here, or browse...",
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

            // 3. Database Manager Group
            GroupBox gbDb = new GroupBox
            {
                Text = "3. Background Database (.txt)",
                Location = new Point(15, 155),
                Size = new Size(615, 60)
            };

            lblDb = new Label
            {
                Text = "Database: Background Database HORIZON.txt (0 items)",
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

            // 4. Preview & Editable Grid
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
                Text = "⚡ AUTO-FILL INTO TALLY BOARD ⚡",
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
                PopulateSheetsFromExcel(loadedFilePath);
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

        private void PopulateSheetsFromExcel(string path)
        {
            try
            {
                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    if (result.Tables.Count > 1)
                    {
                        cbDaySelect.Items.Clear();
                        cbDaySelect.Items.Add("Auto / Default");
                        foreach (DataTable t in result.Tables)
                        {
                            cbDaySelect.Items.Add(t.TableName);
                        }
                        cbDaySelect.SelectedIndex = 0;
                    }
                }
            }
            catch { }
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

        private string FuzzyMatchStation(string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return "";
            string clean = query.Trim().ToLower();

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

        private void ParseCurrentFile()
        {
            if (string.IsNullOrEmpty(loadedFilePath) || !File.Exists(loadedFilePath)) return;

            if (isImageFile)
            {
                ParseImageFile();
            }
            else
            {
                ParseExcelFile();
            }
        }

        private void ParseExcelFile()
        {
            currentSchedule.Clear();
            gridPreview.Rows.Clear();

            bool isNight = rbNight.Checked;
            string selectedDay = cbDaySelect.SelectedItem?.ToString() ?? "Auto / Default";

            try
            {
                using (var stream = File.Open(loadedFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTable table = result.Tables[0];

                    // Check if user picked a specific sheet by name
                    if (result.Tables.Contains(selectedDay))
                    {
                        table = result.Tables[selectedDay];
                    }
                    else if (selectedDay.StartsWith("Day ", StringComparison.OrdinalIgnoreCase))
                    {
                        string targetNum = selectedDay.Substring(4).Trim();
                        foreach (DataTable t in result.Tables)
                        {
                            if (t.TableName.IndexOf(targetNum, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                table = t;
                                break;
                            }
                        }
                    }

                    // Find Day offset if days are stacked on one sheet
                    int offsetRow = 0;
                    if (selectedDay.StartsWith("Day ", StringComparison.OrdinalIgnoreCase))
                    {
                        string dayHeaderKey = "DAY " + selectedDay.Substring(4).Trim();
                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            for (int c = 0; c < Math.Min(table.Columns.Count, 3); c++)
                            {
                                string val = table.Rows[r][c]?.ToString()?.ToUpper() ?? "";
                                if (val.Contains(dayHeaderKey) || val.Contains("TIME-" + dayHeaderKey))
                                {
                                    offsetRow = r;
                                    break;
                                }
                            }
                            if (offsetRow > 0) break;
                        }
                    }

                    int startRow = offsetRow + (isNight ? 7 : 0);
                    int endRow = offsetRow + (isNight ? 12 : 4);

                    for (int c = 2; c < table.Columns.Count; c++)
                    {
                        string colHeader = table.Rows[offsetRow][c]?.ToString() ?? "";
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
            string[] lines = fullText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                Match m = Regex.Match(line, @"\bP(\d+)\b", RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    int pNum = int.Parse(m.Groups[1].Value);
                    List<string> contextLines = new List<string> { line };
                    for (int j = 1; j <= 3 && (i + j) < lines.Length; j++)
                    {
                        contextLines.Add(lines[i + j]);
                    }

                    Tuple<string, string> parsed = DetermineRoleAndStation(pNum, contextLines, isNight);
                    if (!string.IsNullOrEmpty(parsed.Item1) || !string.IsNullOrEmpty(parsed.Item2))
                    {
                        currentSchedule[pNum] = parsed;
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
            // Sync any manual edits made in the preview grid
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

            // Focus Tally Board Window
            foreach (var proc in Process.GetProcesses())
            {
                if (!string.IsNullOrEmpty(proc.MainWindowTitle) &&
                    proc.MainWindowTitle.IndexOf("SCHEDULER TALLY BOARD", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    SetForegroundWindow(proc.MainWindowHandle);
                    break;
                }
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

                // Column 2: Role
                if (!string.IsNullOrEmpty(role))
                {
                    Clipboard.SetText(role);
                    SendKeys.SendWait("^v");
                }
                await Task.Delay(40);
                SendKeys.SendWait("{TAB}");

                // Column 3: Station
                if (!string.IsNullOrEmpty(station))
                {
                    Clipboard.SetText(station);
                    SendKeys.SendWait("^v");
                }
                await Task.Delay(40);
                SendKeys.SendWait("{TAB}");

                // Column 4: Move past dropdown to next row
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
