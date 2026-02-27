using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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

        public MainForm()
        {
            ConfigService.LoadConfigs();
            _adminService = new AdminPrinterService();
            SetupUI();
        }

        private void SetupUI()
        {
            this.Text = "Панель администратора принт-серверов";
            // Срезали высоту окна с 410 до 360
            this.Size = new Size(600, 360);
            this.StartPosition = FormStartPosition.CenterScreen;

            tabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9) };

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
            tabDelete.Controls.Add(gbDelete);

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
        }

        private void SetupStep2Panel()
        {
            pnlStep2.Controls.Add(new Label { Text = "ШАГ 2: Заполнение полей", Top = 20, Left = 20, Width = 500, Font = new Font("Segoe UI", 12, FontStyle.Bold) });
            // Уменьшили отступы между строками для компактности
            int yPos = 65; int spacing = 30;
            int lblWidth = 240; int ctrlLeft = 265; int ctrlWidth = 290;

            pnlStep2.Controls.Add(new Label { Text = "Выберите нужный драйвер:", Top = yPos, Left = 20, Width = lblWidth });
            cmbDriver = new ComboBox { Top = yPos - 3, Left = ctrlLeft, Width = ctrlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            pnlStep2.Controls.Add(cmbDriver);

            yPos += spacing;
            pnlStep2.Controls.Add(new Label { Text = "Филиал расположения(Офис, касса, выдача, менеджеры, склад и т.п.):", Top = yPos, Left = 20, Width = lblWidth });
            cmbBranch = new ComboBox { Top = yPos - 3, Left = ctrlLeft, Width = ctrlWidth, DropDownStyle = ComboBoxStyle.DropDownList };
            if (ConfigService.Cities != null)
                foreach (var city in ConfigService.Cities) cmbBranch.Items.Add(new KeyValuePair<string, string>(city.Code, city.Name));
            cmbBranch.DisplayMember = "Value"; cmbBranch.ValueMember = "Key";
            pnlStep2.Controls.Add(cmbBranch);

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

        private void LogAction(string actionText)
        {
            try
            {
                string logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "admin_log.txt");
                File.AppendAllText(logFile, $"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] {actionText}{Environment.NewLine}");
            }
            catch { }
        }

        private void BtnFindDelete_Click(object sender, EventArgs e)
        {
            string inv = txtDeleteInv.Text.Trim();
            if (string.IsNullOrEmpty(inv)) return;
            lblDeleteInfo.Text = "Поиск на серверах...";
            btnConfirmDelete.Visible = false;
            this.Cursor = Cursors.WaitCursor;
            _foundPrinter = _adminService.FindPrinterForDeletion(inv);
            this.Cursor = Cursors.Default;
            if (_foundPrinter != null)
            {
                lblDeleteInfo.Text = $"Найден: {_foundPrinter.PrinterName}\nСервер: {_foundPrinter.ServerName} | IP: {_foundPrinter.IpAddress}";
                btnConfirmDelete.Visible = true;
            }
            else lblDeleteInfo.Text = "Принтер с таким номером не найден ни на одном сервере.";
        }

        private void BtnConfirmDelete_Click(object sender, EventArgs e)
        {
            if (_foundPrinter == null) return;
            var result = MessageBox.Show($"Удалить принтер '{_foundPrinter.PrinterName}' с сервера {_foundPrinter.ServerName}?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                this.Cursor = Cursors.WaitCursor;
                bool ok = _adminService.DeletePrinterFromServer(_foundPrinter.ServerName, _foundPrinter.PrinterName);
                this.Cursor = Cursors.Default;
                if (ok)
                {
                    LogAction($"УДАЛЕНИЕ: Принтер {_foundPrinter.PrinterName} удален с сервера {_foundPrinter.ServerName}");
                    MessageBox.Show("Принтер успешно удален!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnConfirmDelete.Visible = false;
                    lblDeleteInfo.Text = "Ожидание...";
                    txtDeleteInv.Text = "";
                }
                else MessageBox.Show("Ошибка при удалении.", "Сбой", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void BtnNextStep1_Click(object sender, EventArgs e)
        {
            string ip = txtAddIp.Text.Trim();
            if (string.IsNullOrEmpty(ip)) { MessageBox.Show("Введите IP-адрес!"); return; }

            btnNextStep1.Enabled = false; txtAddIp.Enabled = false;
            lblStatus1.Text = "Проверка IP адреса на предмет дублей на принт-серверах...";
            this.Cursor = Cursors.WaitCursor;

            var checkResult = await Task.Run(() => _adminService.CheckIpOnServers(ip));
            if (checkResult.IpExists && !string.IsNullOrEmpty(checkResult.AttachedPrinterName))
            {
                this.Cursor = Cursors.Default;
                var result = MessageBox.Show($"Принтер '{checkResult.AttachedPrinterName}' с IP {ip} уже существует на {checkResult.ServerName}.\n\nУдалить данное устройство?", "Дубликат", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;
                    lblStatus1.Text = $"Удаление старого принтера...";
                    bool deleted = await Task.Run(() => _adminService.DeletePrinterFromServer(checkResult.ServerName, checkResult.AttachedPrinterName));
                    if (!deleted) { MessageBox.Show("Не удалось удалить старый принтер."); ResetStep1(); return; }
                    LogAction($"УДАЛЕНИЕ ДУБЛЯ IP: {checkResult.AttachedPrinterName} удален с {checkResult.ServerName}");
                }
                else { ResetStep1(); return; }
            }

            _currentIp = ip;
            lblStatus1.Text = "Сбор списка доступных драйверов...";
            await LoadDriversAsync();
            ResetStep1();
            SwitchToStep(2);
        }

        private void ResetStep1() { this.Cursor = Cursors.Default; btnNextStep1.Enabled = true; txtAddIp.Enabled = true; lblStatus1.Text = ""; }

        private async Task LoadDriversAsync()
        {
            cmbDriver.Items.Clear();
            var drivers = await Task.Run(() => _adminService.GetAvailableDrivers());
            if (drivers.Count == 0) cmbDriver.Items.Add("Драйверы не найдены");
            else foreach (var d in drivers) cmbDriver.Items.Add(d);
            if (cmbDriver.Items.Count > 0) cmbDriver.SelectedIndex = 0;
            if (cmbBranch.Items.Count > 0) cmbBranch.SelectedIndex = 0;
            if (cmbType.Items.Count > 0) cmbType.SelectedIndex = 0;
        }

        private async void BtnNextStep2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtLocation.Text) || string.IsNullOrWhiteSpace(txtInvNum.Text))
            { MessageBox.Show("Заполните все поля!"); return; }

            string invNum = txtInvNum.Text.Trim();
            btnNextStep2.Enabled = false; btnBackStep2.Enabled = false;
            lblStatus2.Text = "Проверка инвентарного номера на предмет дублей...";
            this.Cursor = Cursors.WaitCursor;

            var existingInv = await Task.Run(() => _adminService.FindPrinterForDeletion(invNum));
            this.Cursor = Cursors.Default;

            if (existingInv != null)
            {
                var result = MessageBox.Show($"Принтер с таким инв. номером уже найден:\nИмя: {existingInv.PrinterName}\nУдалить старый принтер?", "Дубликат", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor; lblStatus2.Text = "Удаление старого принтера...";
                    bool deleted = await Task.Run(() => _adminService.DeletePrinterFromServer(existingInv.ServerName, existingInv.PrinterName));
                    this.Cursor = Cursors.Default;
                    if (!deleted) { MessageBox.Show("Ошибка удаления."); ResetStep2(); return; }
                    LogAction($"УДАЛЕНИЕ ДУБЛЯ ИНВ: {existingInv.PrinterName} удален с {existingInv.ServerName}");
                }
                else { ResetStep2(); return; }
            }

            ResetStep2();
            PrepareStep3();
            SwitchToStep(3);
        }

        private void ResetStep2() { btnNextStep2.Enabled = true; btnBackStep2.Enabled = true; lblStatus2.Text = ""; this.Cursor = Cursors.Default; }

        private void PrepareStep3()
        {
            string branchCode = ((KeyValuePair<string, string>)cmbBranch.SelectedItem).Key.ToUpper();
            string typeCode = ((KeyValuePair<string, string>)cmbType.SelectedItem).Key;
            string paddedInv = txtInvNum.Text.Trim().PadLeft(6, '0');

            _finalPrinterName = $"{branchCode}{paddedInv}-{typeCode}";
            _finalShareName = paddedInv;
            _finalDriver = cmbDriver.Text;
            _finalLocation = txtLocation.Text.Trim();

            _finalServer = "LG166PS"; // Дефолт
            if (ConfigService.Mappings != null)
            {
                var map = ConfigService.Mappings.Find(m => m.CityCode.Equals(branchCode, StringComparison.OrdinalIgnoreCase));
                if (map != null) _finalServer = map.PrintServer;
            }

            valServer.Text = _finalServer;
            valPrinterName.Text = _finalPrinterName;
            valShareName.Text = _finalShareName;
            valIp.Text = _currentIp;
            valDriver.Text = _finalDriver;
            valLocation.Text = _finalLocation;
        }

        private async void BtnFinish_Click(object sender, EventArgs e)
        {
            if (cmbDriver.Text == "Драйверы не найдены" || string.IsNullOrEmpty(_finalDriver))
            {
                MessageBox.Show("Не выбран драйвер. Невозможно создать принтер.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            btnFinish.Enabled = false; btnBackStep3.Enabled = false;
            lblStatus3.Text = "Создание принтера на сервере. Пожалуйста, подождите...";
            this.Cursor = Cursors.WaitCursor;

            bool created = await Task.Run(() => _adminService.CreatePrinterOnServer(_finalServer, _finalPrinterName, _finalShareName, _currentIp, _finalDriver, _finalLocation));

            this.Cursor = Cursors.Default;

            if (created)
            {
                LogAction($"СОЗДАНИЕ: Принтер {_finalPrinterName} (IP: {_currentIp}) успешно заведен на сервер {_finalServer}");
                MessageBox.Show($"Принтер {_finalPrinterName} успешно заведен на сервер {_finalServer}!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtAddIp.Text = ""; txtInvNum.Text = ""; txtLocation.Text = "";
                SwitchToStep(1);
            }
            else
            {
                MessageBox.Show("Произошла ошибка при создании принтера через WMI. Проверьте права доступа и наличие указанного драйвера на целевом сервере.", "Сбой", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus3.Text = "";
            }

            btnFinish.Enabled = true; btnBackStep3.Enabled = true;
        }
    }
}