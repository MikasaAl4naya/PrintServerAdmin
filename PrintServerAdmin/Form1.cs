using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintServerAdmin
{
    public partial class MainForm : Form
    {
        private AdminPrinterService _adminService;
        private RemotePrinterInfo _foundPrinter;

        private string _currentIp = "";
        private string _finalServer = "";
        private string _finalPrinterName = "";
        private string _finalShareName = "";
        private string _finalDriver = "";
        private string _finalLocation = "";
        private string _suggestedDriver = "";

        // Пути для логирования
        private string _logFilePath = "";
        private string _defaultLocalLogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "admin_log.txt");
        private const string CsvLogHeader = "Timestamp;User;Action";

        private TabControl tabControl;
        private TabPage tabDelete;
        private TabPage tabAdd;

        private TextBox txtDeleteInv;
        private Label lblDeleteInfo;
        private Button btnConfirmDelete;

        private Panel pnlStep1;
        private Panel pnlStep2;
        private Panel pnlStep3;

        private TextBox txtAddIp;
        private Button btnNextStep1;
        private Label lblStatus1;

        private ComboBox cmbDriver;
        private ComboBox cmbBranch;
        private TextBox txtLocation;
        private TextBox txtInvNum;
        private ComboBox cmbType;
        private Label lblSelectedBranch;
        private Button btnNextStep2;
        private Button btnBackStep2;
        private Label lblStatus2;

        private Label valServer;
        private Label valPrinterName;
        private Label valShareName;
        private Label valIp;
        private Label valDriver;
        private Label valLocation;
        private Button btnFinish;
        private Button btnBackStep3;
        private Label lblStatus3;

        private ProgressBar pbDelete;
        private ProgressBar pbStep1;
        private ProgressBar pbStep2;
        private ProgressBar pbStep3;

        public MainForm()
        {
            ConfigService.LoadConfigs();
            LoadLogConfig();

            // Логируем запуск
            LogAction("СИСТЕМА: Программа запущена");

            _adminService = new AdminPrinterService();
            SetupUI();
        }

        private void SetupUI()
        {
            this.Text = "Админ-панель принт-серверов";
            this.Icon = Properties.Resources.printer;
            this.Size = new Size(600, 360);
            this.StartPosition = FormStartPosition.CenterScreen;

            tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9) };

            // === ВКЛАДКА УДАЛЕНИЯ ===
            tabDelete = new TabPage("Удаление принтера");
            GroupBox gbDelete = new GroupBox { Text = "Поиск и удаление", Top = 20, Left = 20, Width = 540, Height = 220 };
            gbDelete.Controls.Add(new Label { Text = "Инв. номер:", Top = 30, Left = 20, Width = 80 });
            txtDeleteInv = new TextBox { Top = 28, Left = 100, Width = 150 };
            gbDelete.Controls.Add(txtDeleteInv);
            Button btnFindDelete = new Button { Text = "Найти на серверах", Top = 26, Left = 260, Width = 150 };
            btnFindDelete.Click += BtnFindDelete_Click;
            gbDelete.Controls.Add(btnFindDelete);
            lblDeleteInfo = new Label { Text = "Введите номер и нажмите найти...", Top = 70, Left = 20, Width = 500, Height = 40, ForeColor = Color.Blue };
            gbDelete.Controls.Add(lblDeleteInfo);
            btnConfirmDelete = new Button { Text = "ПОДТВЕРДИТЬ УДАЛЕНИЕ", Top = 130, Left = 20, Width = 500, Height = 40, BackColor = Color.LightCoral, Visible = false };
            btnConfirmDelete.Click += BtnConfirmDelete_Click;
            gbDelete.Controls.Add(btnConfirmDelete);
            pbDelete = new ProgressBar { Top = 185, Left = 20, Width = 500, Height = 15, Style = ProgressBarStyle.Continuous, Visible = false };
            gbDelete.Controls.Add(pbDelete);
            tabDelete.Controls.Add(gbDelete);

            // === ВКЛАДКА ДОБАВЛЕНИЯ ===
            tabAdd = new TabPage("Добавление принтера");
            pnlStep1 = new Panel { Dock = DockStyle.Fill };
            pnlStep2 = new Panel { Dock = DockStyle.Fill, Visible = false };
            pnlStep3 = new Panel { Dock = DockStyle.Fill, Visible = false };
            SetupStep1Panel();
            SetupStep2Panel();
            SetupStep3Panel();
            tabAdd.Controls.Add(pnlStep1);
            tabAdd.Controls.Add(pnlStep2);
            tabAdd.Controls.Add(pnlStep3);

            tabControl.TabPages.Add(tabDelete);
            tabControl.TabPages.Add(tabAdd);
            this.Controls.Add(tabControl);
        }

        private void LoadLogConfig()
        {
            // По умолчанию ставим локальный путь
            _logFilePath = _defaultLocalLogPath;
            try
            {
                string settingsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json");
                if (File.Exists(settingsFile))
                {
                    string json = File.ReadAllText(settingsFile, Encoding.UTF8);
                    dynamic obj = Newtonsoft.Json.JsonConvert.DeserializeObject(json);
                    if (obj != null && obj.LogFilePath != null)
                    {
                        string targetPath = obj.LogFilePath.ToString();
                        if (!string.IsNullOrWhiteSpace(targetPath))
                        {
                            _logFilePath = targetPath;
                        }
                    }
                }
            }
            catch { }
        }

        private void LogAction(string actionText)
        {
            string userName = Environment.UserName;
            string logEntry = BuildCsvLogEntry(userName, actionText);

            // 1. Пытаемся записать по основному пути (например, сетевому)
            try
            {
                if (WriteToFile(_logFilePath, logEntry)) return;
            }
            catch { }

            // 2. Если основной путь недоступен, пишем локально рядом с .exe
            try
            {
                if (_logFilePath != _defaultLocalLogPath)
                {
                    WriteToFile(_defaultLocalLogPath, logEntry + " (ВНИМАНИЕ: Записано локально, сетевой путь недоступен)");
                }
            }
            catch { }
        }

        private bool WriteToFile(string path, string text)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
            EnsureCsvHeader(path);
            File.AppendAllText(path, text, Encoding.UTF8);
            return true;
        }

        private static string BuildCsvLogEntry(string userName, string actionText)
        {
            string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            return $"{EscapeCsv(ts)};{EscapeCsv(userName)};{EscapeCsv(actionText)}{Environment.NewLine}";
        }

        private static string EscapeCsv(string value)
        {
            if (value == null) value = string.Empty;
            value = value.Replace("\"", "\"\"");
            bool mustQuote = value.Contains(";") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n");
            return mustQuote ? $"\"{value}\"" : value;
        }

        private static void EnsureCsvHeader(string path)
        {
            if (File.Exists(path) && new FileInfo(path).Length > 0) return;
            File.WriteAllText(path, CsvLogHeader + Environment.NewLine, Encoding.UTF8);
        }

        private void SwitchToStep(int stepNumber)
        {
            pnlStep1.Visible = (stepNumber == 1);
            pnlStep2.Visible = (stepNumber == 2);
            pnlStep3.Visible = (stepNumber == 3);
            if (stepNumber == 1) lblStatus1.Text = "";
            if (stepNumber == 2) lblStatus2.Text = "";
            if (stepNumber == 3) lblStatus3.Text = "";
        }

        // --- ШАГ 1: ПРОВЕРКА IP (Исправлено) ---
        private async void BtnNextStep1_Click(object sender, EventArgs e)
        {
            string ip = txtAddIp.Text.Trim();
            if (string.IsNullOrEmpty(ip))
            {
                MessageBox.Show("Введите IP адрес принтера.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!IPAddress.TryParse(ip, out _))
            {
                MessageBox.Show("Некорректный формат IP адреса.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            LogAction($"ШАГ 1: Проверка IP {ip}");
            btnNextStep1.Enabled = false;
            lblStatus1.Text = "Проверка IP на серверах...";
            pbStep1.Value = 0; pbStep1.Visible = true;
            this.Cursor = Cursors.WaitCursor;

            var progress = new Progress<int>(percent => { if (percent >= 0 && percent <= 100) pbStep1.Value = percent; });
            var checkResult = await Task.Run(() => _adminService.CheckIpOnServers(ip, progress));

            this.Cursor = Cursors.Default;
            pbStep1.Visible = false;

            // СЛЕДУЕМ ТЗ:
            // 1. Если нашли ПРИНТЕР (AttachedPrinterName не пустой) — выводим предупреждение
            if (checkResult.IpExists && !string.IsNullOrEmpty(checkResult.AttachedPrinterName))
            {
                LogAction($"ПРЕДУПРЕЖДЕНИЕ: Принтер '{checkResult.AttachedPrinterName}' уже использует IP {ip}");

                var result = MessageBox.Show(
                    $"Принтер '{checkResult.AttachedPrinterName}' с IP адресом '{ip}' уже существует.\n\nУдалить данное устройство?",
                    "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;
                    lblStatus1.Text = "Удаление устройства...";
                    bool delOk = await Task.Run(() => _adminService.DeletePrinterFromServer(checkResult.ServerName, checkResult.AttachedPrinterName));
                    this.Cursor = Cursors.Default;

                    if (delOk) LogAction($"УДАЛЕНИЕ: Устройство {checkResult.AttachedPrinterName} удалено для переустановки");
                    else
                    {
                        LogAction($"ОШИБКА: Не удалось удалить устройство {checkResult.AttachedPrinterName} перед переустановкой");
                        MessageBox.Show("Не удалось удалить существующий принтер. Повторите попытку.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        ResetStep1();
                        return;
                    }
                }
                else
                {
                    LogAction("ОТМЕНА: Пользователь вернулся к вводу IP");
                    ResetStep1();
                    return; // Возвращаемся к вводу IP
                }
            }
            // 2. Если нашли только порт (AttachedPrinterName пустой) — ИГНОРИРУЕМ и идем дальше
            else if (checkResult.IpExists)
            {
                LogAction($"ИНФО: Найден существующий порт {ip} без принтера. Используем его (согласно ТЗ).");
            }

            // Переход к Шагу 2 (выбор драйвера)
            _currentIp = ip;
            _suggestedDriver = await Task.Run(() => _adminService.FindDriverByIpOnServers(ip)) ?? "";
            if (!string.IsNullOrWhiteSpace(_suggestedDriver))
                LogAction($"ИНФО: Найден рекомендуемый драйвер для IP {ip}: {_suggestedDriver}");
            await LoadDriversAsync();
            ResetStep1();
            SwitchToStep(2);
        }

        // --- ФИНАЛ: СОЗДАНИЕ (С исправленным сообщением и логином) ---
        private async void BtnFinish_Click(object sender, EventArgs e)
        {
            LogAction($"ШАГ 3: Попытка создания принтера {_finalPrinterName} (IP: {_currentIp})");
            btnFinish.Enabled = false;
            lblStatus3.Text = "Создание принтера на сервере...";
            pbStep3.Value = 0; pbStep3.Visible = true;
            this.Cursor = Cursors.WaitCursor;

            var progress = new Progress<int>(percent => { if (percent >= 0 && percent <= 100) pbStep3.Value = percent; });
            bool created = await Task.Run(() => _adminService.CreatePrinterOnServer(_finalServer, _finalPrinterName, _finalShareName, _currentIp, _finalDriver, _finalLocation, progress));

            this.Cursor = Cursors.Default; pbStep3.Visible = false;

            if (created)
            {
                LogAction($"УСПЕХ: Принтер {_finalPrinterName} успешно заведен на сервер {_finalServer}");
                MessageBox.Show($"Принтер {_finalPrinterName} успешно заведен на сервер {_finalServer}!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);

                txtAddIp.Text = ""; txtInvNum.Text = ""; txtLocation.Text = "";
                SwitchToStep(1);
            }
            else
            {
                LogAction($"ОШИБКА: Не удалось создать принтер {_finalPrinterName} на сервере {_finalServer}");
                MessageBox.Show("Произошла ошибка при создании принтера через WMI. Проверьте права доступа.", "Сбой", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            btnFinish.Enabled = true;
        }

        // --- ОСТАЛЬНЫЕ МЕТОДЫ (Без изменений) ---
        private void SetupStep1Panel()
        {
            pnlStep1.Controls.Add(new Label { Text = "ШАГ 1: Введите IP адрес принтера", Top = 25, Left = 20, Width = 500, Font = new Font("Segoe UI", 12, FontStyle.Bold) });
            pnlStep1.Controls.Add(new Label { Text = "1. Введите IP адрес принтера:", Top = 80, Left = 20, Width = 200 });
            txtAddIp = new TextBox { Top = 78, Left = 230, Width = 250, Font = new Font("Segoe UI", 10) };
            pnlStep1.Controls.Add(txtAddIp);
            btnNextStep1 = new Button { Text = "Далее >", Top = 120, Left = 230, Width = 150, Height = 35, BackColor = Color.LightGreen };
            btnNextStep1.Click += BtnNextStep1_Click;
            pnlStep1.Controls.Add(btnNextStep1);
            lblStatus1 = new Label { Text = "", Top = 170, Left = 20, Width = 500, Height = 40, ForeColor = Color.DarkBlue, Font = new Font("Segoe UI", 9, FontStyle.Italic) };
            pnlStep1.Controls.Add(lblStatus1);
            pbStep1 = new ProgressBar { Top = 215, Left = 20, Width = 500, Height = 15, Style = ProgressBarStyle.Continuous, Visible = false };
            pnlStep1.Controls.Add(pbStep1);
        }

        private void SetupStep2Panel()
        {
            pnlStep2.Controls.Add(new Label { Text = "ШАГ 2: Заполнение полей", Top = 20, Left = 20, Width = 500, Font = new Font("Segoe UI", 12, FontStyle.Bold) });
            int yPos = 65; int spacing = 30;
            int lblWidth = 240; int ctrlLeft = 265; int ctrlWidth = 290;

            pnlStep2.Controls.Add(new Label { Text = "Выберите нужный драйвер:", Top = yPos, Left = 20, Width = lblWidth });
            cmbDriver = new ComboBox { Top = yPos - 3, Left = ctrlLeft, Width = ctrlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            pnlStep2.Controls.Add(cmbDriver);

            yPos += spacing;
            pnlStep2.Controls.Add(new Label { Text = "Филиал расположения:", Top = yPos, Left = 20, Width = lblWidth });
            cmbBranch = new ComboBox { Top = yPos - 3, Left = ctrlLeft, Width = ctrlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            if (ConfigService.Cities != null)
                foreach (var city in ConfigService.Cities) cmbBranch.Items.Add(new KeyValuePair<string, string>(city.Code, city.Name));
            cmbBranch.DisplayMember = "Value"; cmbBranch.ValueMember = "Key";
            cmbBranch.SelectedIndexChanged += CmbBranch_SelectedIndexChanged;
            pnlStep2.Controls.Add(cmbBranch);

            yPos += spacing;
            lblSelectedBranch = new Label
            {
                Text = "Выбран филиал: не выбран",
                Top = yPos - 2,
                Left = ctrlLeft,
                Width = ctrlWidth,
                Height = 24,
                ForeColor = Color.DarkSlateBlue,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            pnlStep2.Controls.Add(lblSelectedBranch);

            yPos += spacing;
            pnlStep2.Controls.Add(new Label { Text = "Расположение принтера:", Top = yPos, Left = 20, Width = lblWidth });
            txtLocation = new TextBox { Top = yPos - 3, Left = ctrlLeft, Width = ctrlWidth };
            pnlStep2.Controls.Add(txtLocation);

            yPos += spacing;
            pnlStep2.Controls.Add(new Label { Text = "Инвентарный номер принтера:", Top = yPos, Left = 20, Width = lblWidth });
            txtInvNum = new TextBox { Top = yPos - 3, Left = ctrlLeft, Width = 150 };
            txtInvNum.KeyPress += (s, e) => {
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true;
                if (txtInvNum.Text.Length == 0 && e.KeyChar == '0') e.Handled = true;
            };
            pnlStep2.Controls.Add(txtInvNum);

            yPos += spacing;
            pnlStep2.Controls.Add(new Label { Text = "Тип принтера:", Top = yPos, Left = 20, Width = lblWidth });
            cmbType = new ComboBox { Top = yPos - 3, Left = ctrlLeft, Width = ctrlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbType.Items.Add(new KeyValuePair<string, string>("PR", "Обычный принтер (PR)"));
            cmbType.Items.Add(new KeyValuePair<string, string>("CP", "Цветной принтер (CP)"));
            cmbType.Items.Add(new KeyValuePair<string, string>("MF", "Многофункциональное устройство (MF)"));
            cmbType.Items.Add(new KeyValuePair<string, string>("TP", "Термотрансферный принтер (TP)"));
            cmbType.DisplayMember = "Value"; cmbType.ValueMember = "Key";
            pnlStep2.Controls.Add(cmbType);

            yPos += spacing + 15;
            btnBackStep2 = new Button { Text = "< Назад", Top = yPos, Left = 20, Width = 100, Height = 35 };
            btnBackStep2.Click += (s, e) => SwitchToStep(1);
            pnlStep2.Controls.Add(btnBackStep2);

            btnNextStep2 = new Button { Text = "Проверить и продолжить >", Top = yPos, Left = ctrlLeft, Width = ctrlWidth, Height = 35, BackColor = Color.LightGreen };
            btnNextStep2.Click += BtnNextStep2_Click;
            pnlStep2.Controls.Add(btnNextStep2);

            yPos += 40;
            lblStatus2 = new Label { Text = "", Top = yPos, Left = 20, Width = 530, Height = 40, ForeColor = Color.DarkBlue, Font = new Font("Segoe UI", 9, FontStyle.Italic) };
            pnlStep2.Controls.Add(lblStatus2);
            pbStep2 = new ProgressBar { Top = 285, Left = 20, Width = 500, Height = 15, Style = ProgressBarStyle.Continuous, Visible = false };
            pnlStep2.Controls.Add(pbStep2);
        }

        private void SetupStep3Panel()
        {
            pnlStep3.Controls.Add(new Label { Text = "ШАГ 3: Проверка данных", Top = 20, Left = 20, Width = 500, Font = new Font("Segoe UI", 12, FontStyle.Bold) });
            int yPos = 60; int spacing = 25;
            Font boldFont = new Font("Segoe UI", 9, FontStyle.Bold);

            pnlStep3.Controls.Add(new Label { Text = "Сервер:", Top = yPos, Left = 20, Width = 140 });
            valServer = new Label { Text = "-", Top = yPos, Left = 170, Width = 380, Font = boldFont, ForeColor = Color.DarkRed };
            pnlStep3.Controls.Add(valServer);
            yPos += spacing;
            pnlStep3.Controls.Add(new Label { Text = "Имя принтера:", Top = yPos, Left = 20, Width = 140 });
            valPrinterName = new Label { Text = "-", Top = yPos, Left = 170, Width = 380, Font = boldFont };
            pnlStep3.Controls.Add(valPrinterName);
            yPos += spacing;
            pnlStep3.Controls.Add(new Label { Text = "Имя ресурса:", Top = yPos, Left = 20, Width = 140 });
            valShareName = new Label { Text = "-", Top = yPos, Left = 170, Width = 380, Font = boldFont };
            pnlStep3.Controls.Add(valShareName);
            yPos += spacing;
            pnlStep3.Controls.Add(new Label { Text = "IP адрес:", Top = yPos, Left = 20, Width = 140 });
            valIp = new Label { Text = "-", Top = yPos, Left = 170, Width = 380, Font = boldFont };
            pnlStep3.Controls.Add(valIp);
            yPos += spacing;
            pnlStep3.Controls.Add(new Label { Text = "Драйвер:", Top = yPos, Left = 20, Width = 140 });
            valDriver = new Label { Text = "-", Top = yPos, Left = 170, Width = 380, Font = boldFont };
            pnlStep3.Controls.Add(valDriver);
            yPos += spacing;
            pnlStep3.Controls.Add(new Label { Text = "Расположение:", Top = yPos, Left = 20, Width = 140 });
            valLocation = new Label { Text = "-", Top = yPos, Left = 170, Width = 380, Font = boldFont };
            pnlStep3.Controls.Add(valLocation);
            yPos += spacing + 15;

            btnBackStep3 = new Button { Text = "< Назад", Top = yPos, Left = 20, Width = 100, Height = 40 };
            btnBackStep3.Click += (s, e) => SwitchToStep(2);
            pnlStep3.Controls.Add(btnBackStep3);

            btnFinish = new Button { Text = "СОЗДАТЬ ПРИНТЕР", Top = yPos, Left = 170, Width = 380, Height = 40, BackColor = Color.MediumSeaGreen, ForeColor = Color.White, Font = boldFont };
            btnFinish.Click += BtnFinish_Click;
            pnlStep3.Controls.Add(btnFinish);

            yPos += 45;
            lblStatus3 = new Label { Text = "", Top = yPos, Left = 20, Width = 530, Height = 40, ForeColor = Color.DarkBlue, Font = new Font("Segoe UI", 9, FontStyle.Italic) };
            pnlStep3.Controls.Add(lblStatus3);
            pbStep3 = new ProgressBar { Top = 290, Left = 20, Width = 500, Height = 15, Style = ProgressBarStyle.Continuous, Visible = false };
            pnlStep3.Controls.Add(pbStep3);
        }

        private async Task LoadDriversAsync()
        {
            cmbDriver.Items.Clear();
            var progress = new Progress<int>(percent => { if (percent >= 0 && percent <= 100) pbStep1.Value = percent; });
            var drivers = await Task.Run(() => _adminService.GetAvailableDrivers(progress));
            foreach (var d in drivers) cmbDriver.Items.Add(d);
            cmbDriver.SelectedIndex = -1;
            cmbBranch.SelectedIndex = -1;
            if (cmbType.Items.Count > 0) cmbType.SelectedIndex = 0;
            UpdateBranchInfoLabel();

            if (!string.IsNullOrWhiteSpace(_suggestedDriver))
                lblStatus2.Text = $"Рекомендуемый драйвер: {_suggestedDriver}";
            else
                lblStatus2.Text = "Выберите драйвер вручную.";
        }

        private void PrepareStep3()
        {
            string branchCode = ((KeyValuePair<string, string>)cmbBranch.SelectedItem).Key.ToUpper();
            string typeCode = ((KeyValuePair<string, string>)cmbType.SelectedItem).Key;
            string paddedInv = txtInvNum.Text.Trim().PadLeft(6, '0');
            _finalPrinterName = $"{branchCode}{paddedInv}-{typeCode}";
            _finalShareName = paddedInv;
            _finalDriver = cmbDriver.Text;
            _finalLocation = txtLocation.Text.Trim();
            _finalServer = "printsrv01";
            bool hasTypedMap = false;
            if (ConfigService.Mappings != null)
            {
                var typedMap = ConfigService.Mappings.Find(m =>
                    m.CityCode.Equals(branchCode, StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(m.PrinterType) &&
                    m.PrinterType.Equals(typeCode, StringComparison.OrdinalIgnoreCase));
                hasTypedMap = typedMap != null;

                var map = typedMap ?? ConfigService.Mappings.Find(m =>
                    m.CityCode.Equals(branchCode, StringComparison.OrdinalIgnoreCase) &&
                    string.IsNullOrWhiteSpace(m.PrinterType));

                if (map != null) _finalServer = map.PrintServer;
            }

            if (typeCode.Equals("TP", StringComparison.OrdinalIgnoreCase) && !hasTypedMap && ConfigService.TpPrintServers.Count > 0)
                _finalServer = ConfigService.TpPrintServers[0];

            valServer.Text = _finalServer; valPrinterName.Text = _finalPrinterName;
            valShareName.Text = _finalShareName; valIp.Text = _currentIp;
            valDriver.Text = _finalDriver; valLocation.Text = _finalLocation;
        }

        private async void BtnFindDelete_Click(object sender, EventArgs e)
        {
            string inv = txtDeleteInv.Text.Trim();
            if (string.IsNullOrEmpty(inv))
            {
                MessageBox.Show("Введите инвентарный номер.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            LogAction($"ПОИСК (УДАЛЕНИЕ): Инв. № {inv}");
            lblDeleteInfo.Text = "Поиск на серверах...";
            btnConfirmDelete.Visible = false;
            pbDelete.Value = 0; pbDelete.Visible = true;
            this.Cursor = Cursors.WaitCursor;
            var progress = new Progress<int>(percent => { if (percent >= 0 && percent <= 100) pbDelete.Value = percent; });
            _foundPrinter = await Task.Run(() => _adminService.FindPrinterForDeletion(inv, progress));
            this.Cursor = Cursors.Default;
            pbDelete.Visible = false;
            if (_foundPrinter != null)
            {
                lblDeleteInfo.Text = $"Найден: {_foundPrinter.PrinterName}\nIP: {_foundPrinter.IpAddress}\nСервер: {_foundPrinter.ServerName}";
                LogAction($"РЕЗУЛЬТАТ: На сервере {_foundPrinter.ServerName} найден {_foundPrinter.PrinterName}");
                btnConfirmDelete.Visible = true;
            }
            else { lblDeleteInfo.Text = "Принтер не найден."; LogAction($"РЕЗУЛЬТАТ: Принтер № {inv} не найден"); }
        }

        private async void BtnConfirmDelete_Click(object sender, EventArgs e)
        {
            if (_foundPrinter == null) return;
            var result = MessageBox.Show($"Удалить принтер {_foundPrinter.PrinterName} (IP: {_foundPrinter.IpAddress})?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                this.Cursor = Cursors.WaitCursor;
                bool ok = await Task.Run(() => _adminService.DeletePrinterFromServer(_foundPrinter.ServerName, _foundPrinter.PrinterName));
                this.Cursor = Cursors.Default;
                if (ok)
                {
                    LogAction($"УСПЕХ: Принтер {_foundPrinter.PrinterName} удален с {_foundPrinter.ServerName}");
                    MessageBox.Show("Успешно!");
                    btnConfirmDelete.Visible = false;
                    txtDeleteInv.Text = "";
                }
                else LogAction($"ОШИБКА: Не удалось удалить {_foundPrinter.PrinterName} с {_foundPrinter.ServerName}");
            }
            else LogAction($"ОТМЕНА: Сотрудник отказался от удаления {_foundPrinter.PrinterName}");
        }

        private async void BtnNextStep2_Click(object sender, EventArgs e)
        {
            if (cmbDriver.SelectedItem == null)
            {
                MessageBox.Show("Драйвер не выбран. Выберите драйвер из списка.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbBranch.SelectedItem == null)
            {
                MessageBox.Show("Филиал не выбран. Выберите филиал из списка.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbType.SelectedItem == null)
            {
                MessageBox.Show("Тип принтера не выбран.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtLocation.Text) || string.IsNullOrWhiteSpace(txtInvNum.Text))
            {
                MessageBox.Show("Заполните все поля шага 2.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string invNum = txtInvNum.Text.Trim();
            LogAction($"ШАГ 2: Проверка инв. № {invNum}");
            btnNextStep2.Enabled = false;
            lblStatus2.Text = "Проверка инвентарника...";
            pbStep2.Value = 0; pbStep2.Visible = true;
            this.Cursor = Cursors.WaitCursor;
            var progress = new Progress<int>(percent => { if (percent >= 0 && percent <= 100) pbStep2.Value = percent; });
            var existingInv = await Task.Run(() => _adminService.FindPrinterForDeletion(invNum, progress));
            this.Cursor = Cursors.Default; pbStep2.Visible = false;
            if (existingInv != null)
            {
                LogAction($"ПРЕДУПРЕЖДЕНИЕ: Инв. № {invNum} занят на {existingInv.ServerName}");
                var result = MessageBox.Show(
                    $"Найден принтер с таким инвентарным номером.\n\nИмя: {existingInv.PrinterName}\nИнв. номер: {invNum}\nСервер: {existingInv.ServerName}\n\nУдалить его?",
                    "Дубликат инвентарного номера",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;
                    bool delOk = await Task.Run(() => _adminService.DeletePrinterFromServer(existingInv.ServerName, existingInv.PrinterName));
                    this.Cursor = Cursors.Default;
                    if (delOk) LogAction($"УДАЛЕНИЕ ДУБЛЯ (ИНВ): {existingInv.PrinterName} удален");
                    else
                    {
                        LogAction($"ОШИБКА: Не удалось удалить дубль инвентарного номера {existingInv.PrinterName}");
                        MessageBox.Show("Не удалось удалить найденный дубль. Процесс остановлен.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        ResetStep2();
                        return;
                    }
                }
                else { LogAction("ОТМЕНА: Установка прервана из-за занятого инвентарника"); ResetStep2(); return; }
            }
            ResetStep2();
            PrepareStep3();
            SwitchToStep(3);
        }

        private void ResetStep1() { this.Cursor = Cursors.Default; btnNextStep1.Enabled = true; lblStatus1.Text = ""; pbStep1.Visible = false; }
        private void ResetStep2() { btnNextStep2.Enabled = true; lblStatus2.Text = ""; pbStep2.Visible = false; }

        private void CmbBranch_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateBranchInfoLabel();
        }

        private void UpdateBranchInfoLabel()
        {
            if (lblSelectedBranch == null) return;

            if (cmbBranch?.SelectedItem is KeyValuePair<string, string> selected)
                lblSelectedBranch.Text = $"Выбран филиал: {selected.Value} ({selected.Key})";
            else
                lblSelectedBranch.Text = "Выбран филиал: не выбран";
        }
    }
}