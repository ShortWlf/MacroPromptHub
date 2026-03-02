using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MacroPromptHub
{
    public partial class Form1 : Form
    {
        private readonly string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "prompts.txt");
        private bool startWithWindows = false;
        private bool minimizeToTray = true;
        private bool stayOnTop = false;
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private Dictionary<int, string> hotkeyMap = new Dictionary<int, string>();
        private int hotkeyCounter = 1000;
        private bool hotkeyAssignMode = false;
        private Button currentHoverButton = null;
        private IntPtr lastBrowserHandle = IntPtr.Zero;
        private string lastBrowserUrl = "";

        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true;
            this.KeyDown += Form1_KeyDown;
            LoadSettingsFromConfig();
            this.TopMost = stayOnTop;
            if (menuStrip1 != null) menuStrip1.Dock = DockStyle.Top;
            if (flowLayoutPanel1 != null)
            {
                flowLayoutPanel1.Dock = DockStyle.Fill;
                flowLayoutPanel1.AutoScroll = true;
                flowLayoutPanel1.BringToFront();
            }
            SetupDynamicMenu();
            SetupTrayIcon();
            trayIcon.DoubleClick += (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; this.BringToFront(); this.Activate(); };
            this.Resize += Form1_Resize;
            this.FormClosing += Form1_FormClosing;
            LoadMacroButtons();
            RegisterHotkeysFromConfig();
            if (startWithWindows) SetStartup(true);
        }

        private void ExecuteMacro(string target, string content, string command)
        {
            string t = target.ToLower();
            if (t == "browser")
            {
                string prompt = "";
                if (command.StartsWith("prompt:")) prompt = command.Substring("prompt:".Length).Trim();
                ExecuteBrowserMacro(content, prompt);
                return;
            }
            if (t == "focus") { ExecuteFocusMacro(content, command); return; }
            if (t == "run") { ExecuteRunMacro(content); return; }
            if (t == "continue") { ExecuteContinueSessionMacro(content); return; }

            ExecuteAppMacro(target, content, command);
        }

        private void ExecuteAppMacro(string target, string content, string command)
        {
            Task.Run(() =>
            {
                try
                {
                    content = content.Replace("\\n", "\n");
                    this.Invoke(new Action(() => Clipboard.SetText(content)));

                    string procName = target.Replace(".exe", "").Trim().ToLowerInvariant();
                    Process existing = Process.GetProcessesByName(procName).FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);

                    Process targetProc = existing;
                    bool launchedByHub = false;

                    if (targetProc == null)
                    {
                        if (procName == "cmd") targetProc = Process.Start("cmd.exe", "/K");
                        else targetProc = Process.Start(procName + ".exe");
                        launchedByHub = true;
                    }

                    IntPtr hwnd = WaitForMainWindow(targetProc);
                    if (hwnd == IntPtr.Zero) return;

                    ForceForegroundNuclear(hwnd);
                    int focusCheck = 0;
                    while (GetForegroundWindow() != hwnd && focusCheck < 50)
                    {
                        ForceForegroundNuclear(hwnd);
                        Thread.Sleep(100);
                        focusCheck++;
                    }

                    Thread.Sleep(launchedByHub ? 1000 : 350);
                    SendKeys.SendWait("^v");
                    if (command.Contains("enter")) { Thread.Sleep(150); SendKeys.SendWait("{ENTER}"); }
                }
                catch (Exception ex) { this.Invoke(new Action(() => MessageBox.Show("Macro Error: " + ex.Message))); }
            });
        }

        private void ExecuteFocusMacro(string content, string command)
        {
            Task.Run(() =>
            {
                try
                {
                    content = content.Replace("\\n", "\n");
                    this.Invoke(new Action(() => Clipboard.SetText(content)));
                    Thread.Sleep(200);
                    SendKeys.SendWait("^v");
                    if (command.Contains("enter")) { Thread.Sleep(100); SendKeys.SendWait("{ENTER}"); }
                }
                catch (Exception ex) { this.Invoke(new Action(() => MessageBox.Show("Focus Macro Error: " + ex.Message))); }
            });
        }

        private void ExecuteRunMacro(string appPath)
        {
            Task.Run(() =>
            {
                try
                {
                    string path = appPath.Trim();
                    if (File.Exists(path) || Directory.Exists(path))
                    {
                        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                        return;
                    }
                    Process existing = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(path)).FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);
                    if (existing != null)
                    {
                        IntPtr hwnd = existing.MainWindowHandle;
                        if (hwnd != IntPtr.Zero) ForceForegroundNuclear(hwnd);
                        return;
                    }
                    Process.Start(path);
                }
                catch (Exception ex) { this.Invoke(new Action(() => MessageBox.Show("Run Macro Error: " + ex.Message))); }
            });
        }

        private void ExecuteContinueSessionMacro(string prompt)
        {
            Task.Run(() =>
            {
                try
                {
                    prompt = prompt.Replace("\\n", "\n");
                    Process browser = null;
                    IntPtr hwnd = IntPtr.Zero;
                    if (lastBrowserHandle != IntPtr.Zero && IsWindow(lastBrowserHandle))
                    {
                        hwnd = lastBrowserHandle;
                        browser = Process.GetProcesses().FirstOrDefault(p => p.MainWindowHandle == hwnd);
                    }
                    if (browser == null)
                    {
                        browser = LaunchUrlAndGetBrowser("https://copilot.microsoft.com/?newchat");
                        if (browser == null) return;
                        hwnd = WaitForMainWindow(browser);
                        lastBrowserHandle = hwnd;
                        lastBrowserUrl = "https://copilot.microsoft.com/?newchat";
                    }
                    ForceForegroundNuclear(hwnd);
                    int focusCheck = 0;
                    while (GetForegroundWindow() != hwnd && focusCheck < 40)
                    {
                        ForceForegroundNuclear(hwnd);
                        Thread.Sleep(100);
                        focusCheck++;
                    }
                    Thread.Sleep(1500);
                    if (!string.IsNullOrWhiteSpace(prompt))
                    {
                        this.Invoke(new Action(() => Clipboard.SetText(prompt)));
                        SendKeys.SendWait("^v");
                        Thread.Sleep(150);
                        SendKeys.SendWait("{ENTER}");
                    }
                }
                catch (Exception ex) { this.Invoke(new Action(() => MessageBox.Show("Continue Session Error: " + ex.Message))); }
            });
        }

        private void ExecuteBrowserMacro(string url, string prompt)
        {
            Task.Run(() =>
            {
                try
                {
                    prompt = prompt.Replace("\\n", "\n");
                    Process browser = LaunchUrlAndGetBrowser(url);
                    if (browser == null)
                    {
                        this.Invoke(new Action(() => MessageBox.Show("Could not detect browser window.")));
                        return;
                    }
                    IntPtr hwnd = WaitForMainWindow(browser);
                    if (hwnd == IntPtr.Zero) return;
                    lastBrowserHandle = hwnd;
                    lastBrowserUrl = url;
                    ForceForegroundNuclear(hwnd);
                    int focusCheck = 0;
                    while (GetForegroundWindow() != hwnd && focusCheck < 40)
                    {
                        ForceForegroundNuclear(hwnd);
                        Thread.Sleep(100);
                        focusCheck++;
                    }
                    Thread.Sleep(2000);
                    if (!string.IsNullOrWhiteSpace(prompt))
                    {
                        this.Invoke(new Action(() => Clipboard.SetText(prompt)));
                        SendKeys.SendWait("^v");
                        Thread.Sleep(150);
                        SendKeys.SendWait("{ENTER}");
                    }
                }
                catch (Exception ex) { this.Invoke(new Action(() => MessageBox.Show("Browser Macro Error: " + ex.Message))); }
            });
        }

        private Process LaunchUrlAndGetBrowser(string url)
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            string[] browsers = { "msedge", "chrome", "firefox", "brave", "opera", "vivaldi" };
            Process found = null;
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < 8000)
            {
                foreach (var b in browsers)
                {
                    found = Process.GetProcessesByName(b).FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);
                    if (found != null) return found;
                }
                Thread.Sleep(100);
            }
            return null;
        }

        private IntPtr WaitForMainWindow(Process proc, int timeoutMs = 30000)
        {
            var sw = Stopwatch.StartNew();
            IntPtr handle = IntPtr.Zero;
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                proc.Refresh();
                handle = proc.MainWindowHandle;
                if (handle != IntPtr.Zero) return handle;
                var sameName = Process.GetProcessesByName(proc.ProcessName).FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);
                if (sameName != null) return sameName.MainWindowHandle;
                Thread.Sleep(50);
            }
            return IntPtr.Zero;
        }

        private void ForceForegroundNuclear(IntPtr hWnd)
        {
            if (IsIconic(hWnd)) ShowWindow(hWnd, SW_RESTORE);
            uint targetThread = GetWindowThreadProcessId(hWnd, IntPtr.Zero);
            uint currentThread = GetCurrentThreadId();
            if (targetThread != currentThread)
            {
                AttachThreadInput(currentThread, targetThread, true);
                SetForegroundWindow(hWnd);
                SetFocus(hWnd);
                AttachThreadInput(currentThread, targetThread, false);
            }
            else SetForegroundWindow(hWnd);
        }

        private void RegisterHotkeysFromConfig()
        {
            UnregisterAllHotkeys();
            hotkeyMap.Clear();
            if (!File.Exists(configPath)) return;
            var lines = File.ReadAllLines(configPath);
            foreach (var line in lines)
            {
                if (!line.StartsWith("hotkey:")) continue;
                var parts = line.Substring(7).Split('|');
                if (parts.Length != 2) continue;
                string hotkey = parts[0].Trim();
                string macroName = parts[1].Trim();
                int modifiers = 0;
                Keys key = Keys.None;
                if (hotkey.Contains("CTRL")) modifiers |= 0x2;
                if (hotkey.Contains("ALT")) modifiers |= 0x1;
                if (hotkey.Contains("SHIFT")) modifiers |= 0x4;
                string last = hotkey.Split('+').Last();
                key = (Keys)Enum.Parse(typeof(Keys), last);
                int id = hotkeyCounter++;
                RegisterHotKey(this.Handle, id, modifiers, (int)key);
                hotkeyMap[id] = macroName;
            }
        }

        private void UnregisterAllHotkeys()
        {
            foreach (var kvp in hotkeyMap.ToList()) UnregisterHotKey(this.Handle, kvp.Key);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312)
            {
                int id = m.WParam.ToInt32();
                if (hotkeyMap.ContainsKey(id)) TriggerMacroByName(hotkeyMap[id]);
            }
            base.WndProc(ref m);
        }

        private void TriggerMacroByName(string name)
        {
            foreach (Control c in flowLayoutPanel1.Controls)
            {
                Button btn = c as Button;
                if (btn != null && btn.Text == name) { btn.PerformClick(); return; }
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (!hotkeyAssignMode || currentHoverButton == null) return;
            if (e.KeyCode == Keys.Escape)
            {
                RemoveHotkeyForButton(currentHoverButton.Text);
                RegisterHotkeysFromConfig();
                e.Handled = true;
                return;
            }
            if (e.KeyCode == Keys.ControlKey || e.KeyCode == Keys.ShiftKey || e.KeyCode == Keys.Menu) return;
            string hotkey = "";
            if (e.Control) hotkey += "CTRL+";
            if (e.Alt) hotkey += "ALT+";
            if (e.Shift) hotkey += "SHIFT+";
            hotkey += e.KeyCode.ToString();
            AddOrUpdateHotkeyInConfig(hotkey, currentHoverButton.Text);
            RegisterHotkeysFromConfig();
            e.Handled = true;
        }

        private void AddOrUpdateHotkeyInConfig(string hotkey, string buttonName)
        {
            var lines = File.ReadAllLines(configPath).ToList();
            lines.RemoveAll(l => l.StartsWith("hotkey:") && l.EndsWith("|" + buttonName));
            lines.Add("hotkey:" + hotkey + "|" + buttonName);
            File.WriteAllLines(configPath, lines);
        }

        private void RemoveHotkeyForButton(string buttonName)
        {
            var lines = File.ReadAllLines(configPath).ToList();
            lines.RemoveAll(l => l.StartsWith("hotkey:") && l.EndsWith("|" + buttonName));
            File.WriteAllLines(configPath, lines);
        }

        private void SetupDynamicMenu()
        {
            if (menuStrip1 == null) return;
            menuStrip1.Items.Clear();
            var fileMenu = new ToolStripMenuItem("File");
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("Edit Config", null, EditConfigToolStripMenuItem_Click));
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("Reset Config", null, ResetConfigToolStripMenuItem_Click));
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("Reload Config", null, ReloadConfigToolStripMenuItem_Click));
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("Exit", null, ExitToolStripMenuItem_Click));

            var optionsMenu = new ToolStripMenuItem("Options");
            var startWin = new ToolStripMenuItem("Start with Windows") { Checked = startWithWindows };
            startWin.Click += (s, e) => { startWithWindows = !startWithWindows; startWin.Checked = startWithWindows; SetStartup(startWithWindows); UpdateConfigSetting("start_with_windows", startWithWindows); };
            var minTray = new ToolStripMenuItem("Minimize to Tray") { Checked = minimizeToTray };
            minTray.Click += (s, e) => { minimizeToTray = !minimizeToTray; minTray.Checked = minimizeToTray; UpdateConfigSetting("minimize_to_tray", minimizeToTray); };
            var stayTopItem = new ToolStripMenuItem("Stay on Top") { Checked = stayOnTop };
            stayTopItem.Click += (s, e) => { stayOnTop = !stayOnTop; stayTopItem.Checked = stayOnTop; this.TopMost = stayOnTop; UpdateConfigSetting("stay_on_top", stayOnTop); };
            var hotkeyAssignItem = new ToolStripMenuItem("Hotkey Assignment Mode") { Checked = hotkeyAssignMode };
            hotkeyAssignItem.Click += (s, e) => { hotkeyAssignMode = !hotkeyAssignMode; hotkeyAssignItem.Checked = hotkeyAssignMode; };
            optionsMenu.DropDownItems.Add(startWin);
            optionsMenu.DropDownItems.Add(minTray);
            optionsMenu.DropDownItems.Add(stayTopItem);
            optionsMenu.DropDownItems.Add(new ToolStripSeparator());
            optionsMenu.DropDownItems.Add(hotkeyAssignItem);
            var platformsMenu = new ToolStripMenuItem("Platforms");
            platformsMenu.DropDownItems.Add(new ToolStripMenuItem("Copilot", null, (s, e) => LaunchUrlAndGetBrowser("https://copilot.microsoft.com/?newchat")));
            platformsMenu.DropDownItems.Add(new ToolStripMenuItem("Gemini", null, (s, e) => LaunchUrlAndGetBrowser("https://gemini.google.com/app")));
            platformsMenu.DropDownItems.Add(new ToolStripMenuItem("Grok", null, (s, e) => LaunchUrlAndGetBrowser("https://grok.x.ai")));
            platformsMenu.DropDownItems.Add(new ToolStripMenuItem("ChatGPT", null, (s, e) => LaunchUrlAndGetBrowser("https://chat.openai.com")));
            platformsMenu.DropDownItems.Add(new ToolStripMenuItem("Midjourney", null, (s, e) => LaunchUrlAndGetBrowser("https://www.midjourney.com")));
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileMenu, optionsMenu, platformsMenu });
        }

        private void LoadSettingsFromConfig()
        {
            if (!File.Exists(configPath)) return;
            try
            {
                var lines = File.ReadAllLines(configPath);
                foreach (var line in lines)
                {
                    if (line.StartsWith("settings:start_with_windows:")) startWithWindows = line.EndsWith("true");
                    else if (line.StartsWith("settings:minimize_to_tray:")) minimizeToTray = line.EndsWith("true");
                    else if (line.StartsWith("settings:stay_on_top:")) stayOnTop = line.EndsWith("true");
                }
            }
            catch { }
        }

        private void LoadMacroButtons()
        {
            if (flowLayoutPanel1 == null) return;
            flowLayoutPanel1.Controls.Clear();
            if (!File.Exists(configPath)) CreateDefaultConfig();

            string currentTarget = "browser";
            foreach (var line in File.ReadAllLines(configPath))
            {
                if (line.StartsWith("target:")) currentTarget = line.Split(':')[1].Trim();
                else if (line.Contains("|") && !line.StartsWith("hotkey:") && !line.StartsWith("settings:") && !line.TrimStart().StartsWith("#"))
                {
                    var p = line.Split('|');
                    var btn = new Button { Text = p[0].Trim(), Width = 215, Height = 45, Margin = new Padding(5), FlatStyle = FlatStyle.Flat };
                    string content = p.Length > 1 ? p[1].Trim() : "";
                    string cmd = p.Length > 2 ? p[2].Trim() : "";
                    string finalTarget = currentTarget;

                    btn.Click += (s, e) => ExecuteMacro(finalTarget, content, cmd);
                    btn.MouseEnter += (s, e) => currentHoverButton = btn;
                    btn.MouseLeave += (s, e) => { if (currentHoverButton == btn) currentHoverButton = null; };
                    btn.MouseUp += Button_MouseUp;
                    flowLayoutPanel1.Controls.Add(btn);
                }
            }
            flowLayoutPanel1.MouseUp += FlowLayoutPanel1_MouseUp;
        }

        private void Button_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            Button btn = sender as Button;
            if (btn == null) return;
            var menu = new ContextMenuStrip();
            menu.Items.Add("Edit Macro", null, (s, ev) => EditMacro(btn.Text));
            menu.Items.Add("Delete Macro", null, (s, ev) => DeleteMacro(btn.Text));
            menu.Show(Cursor.Position);
        }

        private void FlowLayoutPanel1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            var menu = new ContextMenuStrip();
            menu.Items.Add("Add New Macro", null, (s, ev) => AddNewMacro());
            menu.Items.Add("Reload Config", null, (s, ev) => { LoadMacroButtons(); RegisterHotkeysFromConfig(); });
            menu.Items.Add("Edit Config", null, EditConfigToolStripMenuItem_Click);
            menu.Show(Cursor.Position);
        }

        private void DeleteMacro(string buttonName)
        {
            var lines = File.ReadAllLines(configPath).ToList();
            int index = lines.FindIndex(l => l.StartsWith(buttonName + "|"));
            if (index >= 0)
            {
                lines.RemoveAt(index);
                File.WriteAllLines(configPath, lines);
                LoadMacroButtons();
                RegisterHotkeysFromConfig();
            }
        }

        // --- FIXED EDIT MACRO TO REMEMBER WINDOW SELECTION ---
        private void EditMacro(string buttonName)
        {
            var lines = File.ReadAllLines(configPath).ToList();
            int index = lines.FindIndex(l => l.StartsWith(buttonName + "|"));
            if (index < 0) return;

            string line = lines[index];
            string[] parts = line.Split('|');
            string oldName = parts[0];
            string oldContent = parts.Length > 1 ? parts[1] : "";
            string oldCommand = parts.Length > 2 ? parts[2] : "";

            string target = "browser";
            for (int i = index - 1; i >= 0; i--)
            {
                if (lines[i].StartsWith("target:")) { target = lines[i].Substring(7).Trim(); break; }
            }

            using (var dlg = new AddMacroForm())
            {
                dlg.SetValues(oldName, target, oldContent, oldCommand);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    // PERSISTENCE FIX: 
                    // 1. Remove the old button line
                    lines.RemoveAt(index);

                    // 2. Remove the old target header immediately above it so it doesn't duplicate
                    if (index > 0 && lines[index - 1].StartsWith("target:"))
                    {
                        lines.RemoveAt(index - 1);
                        index--; // adjust insertion point
                    }

                    // 3. Write the specific new target and updated button info
                    lines.Insert(index, "target:" + dlg.Target);
                    lines.Insert(index + 1, dlg.MacroName + "|" + dlg.Content + "|" + dlg.Command);

                    File.WriteAllLines(configPath, lines);
                    LoadMacroButtons();
                    RegisterHotkeysFromConfig();
                }
            }
        }

        private void AddNewMacro()
        {
            using (var dlg = new AddMacroForm())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    if (string.IsNullOrWhiteSpace(dlg.MacroName) || string.IsNullOrWhiteSpace(dlg.Target)) return;
                    var lines = File.ReadAllLines(configPath).ToList();

                    // Always add the target header so the button remembers its window
                    lines.Add("target:" + dlg.Target);
                    lines.Add(dlg.MacroName + "|" + dlg.Content + "|" + dlg.Command);

                    File.WriteAllLines(configPath, lines);
                    LoadMacroButtons();
                    RegisterHotkeysFromConfig();
                }
            }
        }

        private void CreateDefaultConfig()
        {
            File.WriteAllText(configPath,
                "settings:start_with_windows:false\nsettings:minimize_to_tray:true\nsettings:stay_on_top:false\n\n" +
                "target:notepad\nNotepad Hello|Hello!|enter\n\ntarget:run\nLaunch Notepad|notepad.exe\n\ntarget:focus\nFocus Test|text|enter\n"
            );
        }

        private void SetupTrayIcon()
        {
            trayIcon = new NotifyIcon { Icon = SystemIcons.Application, Visible = false, Text = "Macro Prompt Hub" };
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Show App", null, (s, e) => { this.Show(); this.WindowState = FormWindowState.Normal; });
            trayMenu.Items.Add("Exit", null, ExitToolStripMenuItem_Click);
            trayIcon.ContextMenuStrip = trayMenu;
        }

        private void Form1_Resize(object sender, EventArgs e) { if (minimizeToTray && WindowState == FormWindowState.Minimized) { this.Hide(); trayIcon.Visible = true; } }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e) { if (minimizeToTray && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; this.Hide(); trayIcon.Visible = true; } else UnregisterAllHotkeys(); }
        private void EditConfigToolStripMenuItem_Click(object sender, EventArgs e) { Process.Start(new ProcessStartInfo("notepad.exe", "\"" + configPath + "\"") { UseShellExecute = true }); }
        private void ReloadConfigToolStripMenuItem_Click(object sender, EventArgs e) { LoadMacroButtons(); RegisterHotkeysFromConfig(); }
        private void ResetConfigToolStripMenuItem_Click(object sender, EventArgs e) { CreateDefaultConfig(); LoadMacroButtons(); RegisterHotkeysFromConfig(); }
        private void ExitToolStripMenuItem_Click(object sender, EventArgs e) { UnregisterAllHotkeys(); Application.Exit(); }

        private void UpdateConfigSetting(string k, bool v)
        {
            try { var lines = File.ReadAllLines(configPath).ToList(); lines.RemoveAll(l => l.StartsWith("settings:" + k + ":")); lines.Insert(0, "settings:" + k + ":" + v.ToString().ToLower()); File.WriteAllLines(configPath, lines); } catch { }
        }

        private void SetStartup(bool enable)
        {
            try { using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true)) { if (enable) key.SetValue("MacroPromptHub", Application.ExecutablePath); else key.DeleteValue("MacroPromptHub", false); } } catch { }
        }

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hWnd);
        private const int SW_RESTORE = 9;
    }

    public class AddMacroForm : Form
    {
        public string MacroName => txtName.Text.Trim();
        public string Target => cmbTarget.Text.Trim();
        public string Content => txtContent.Text.Trim();
        public string Command => txtCommand.Text.Trim();

        TextBox txtName = new TextBox();
        ComboBox cmbTarget = new ComboBox();
        ComboBox cmbWindowPicker = new ComboBox();
        Button btnRefresh = new Button();
        TextBox txtContent = new TextBox();
        TextBox txtCommand = new TextBox();
        Button btnOk = new Button();
        Button btnCancel = new Button();

        public AddMacroForm()
        {
            this.Text = "Add / Edit Macro";
            this.Size = new Size(420, 380);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            AddControl(new Label { Text = "Button Name:", Left = 10, Top = 20, Width = 120 });
            txtName = (TextBox)AddControl(new TextBox { Left = 140, Top = 20, Width = 220 });

            AddControl(new Label { Text = "Target Active Window:", Left = 10, Top = 60, Width = 130 });
            cmbWindowPicker = (ComboBox)AddControl(new ComboBox { Left = 140, Top = 60, Width = 180, DropDownStyle = ComboBoxStyle.DropDownList });
            btnRefresh = (Button)AddControl(new Button { Text = "🔄", Left = 325, Top = 58, Width = 35 });
            btnRefresh.Click += (s, e) => RefreshWindowList();

            AddControl(new Label { Text = "Target Process:", Left = 10, Top = 100, Width = 120 });
            cmbTarget = (ComboBox)AddControl(new ComboBox { Left = 140, Top = 100, Width = 220 });
            cmbTarget.Items.AddRange(new string[] { "notepad", "cmd", "browser", "run", "focus", "continue" });

            cmbWindowPicker.SelectedIndexChanged += (s, e) => { if (cmbWindowPicker.SelectedItem is WindowItem item) cmbTarget.Text = item.ProcessName; };

            AddControl(new Label { Text = "Content:", Left = 10, Top = 140, Width = 120 });
            txtContent = (TextBox)AddControl(new TextBox { Left = 140, Top = 140, Width = 220 });

            AddControl(new Label { Text = "Command:", Left = 10, Top = 180, Width = 120 });
            txtCommand = (TextBox)AddControl(new TextBox { Left = 140, Top = 180, Width = 220 });

            btnOk = (Button)AddControl(new Button { Text = "OK", Left = 140, Top = 260, Width = 80, DialogResult = DialogResult.OK });
            btnCancel = (Button)AddControl(new Button { Text = "Cancel", Left = 240, Top = 260, Width = 80, DialogResult = DialogResult.Cancel });
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
            RefreshWindowList();
        }

        private Control AddControl(Control c) { this.Controls.Add(c); return c; }

        private void RefreshWindowList()
        {
            cmbWindowPicker.Items.Clear();
            var processes = Process.GetProcesses().Where(p => !string.IsNullOrEmpty(p.MainWindowTitle)).ToList();
            foreach (var p in processes) cmbWindowPicker.Items.Add(new WindowItem { Title = p.MainWindowTitle, ProcessName = p.ProcessName });
        }

        public void SetValues(string name, string target, string content, string command)
        {
            txtName.Text = name;
            cmbTarget.Text = target;
            txtContent.Text = content;
            txtCommand.Text = command;
        }
    }

    public class WindowItem
    {
        public string Title { get; set; }
        public string ProcessName { get; set; }
        public override string ToString() => Title;
    }
}