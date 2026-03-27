using System;
using System.Management;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PrintServerAdmin
{
    public class RemotePrinterInfo
    {
        public string ServerName { get; set; }
        public string PrinterName { get; set; }
        public string IpAddress { get; set; }
    }

    public class IpCheckResult
    {
        public bool IpExists { get; set; }
        public string PortName { get; set; }
        public string ServerName { get; set; }
        public string AttachedPrinterName { get; set; }
    }

    public class AdminPrinterService
    {
        // === СУПЕРБЫСТРЫЙ ПОИСК ПО ИНВЕНТАРНИКУ ===
        public RemotePrinterInfo FindPrinterForDeletion(string invNumber, IProgress<int> progress = null)
        {
            string paddedInv = invNumber.PadLeft(6, '0');
            var servers = ConfigService.PrintServers;
            int totalServers = servers.Count;
            if (totalServers == 0) return null;

            int step = 0;
            // 4 шага на каждый сервер для плавного прогресс-бара
            int totalSteps = totalServers * 4;
            void Report() { step++; progress?.Report((step * 100) / totalSteps); }

            progress?.Report(0);

            foreach (var server in servers)
            {
                try
                {
                    ConnectionOptions options = new ConnectionOptions { Timeout = TimeSpan.FromSeconds(3) };
                    ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2", options);
                    scope.Connect();
                    Report(); // Шаг 1: Подключились

                    // Махом скачиваем все порты в кэш (Секрет скорости!)
                    var portDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    using (var portSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, HostAddress FROM Win32_TCPIPPrinterPort")))
                    {
                        foreach (ManagementObject port in portSearcher.Get())
                        {
                            string pName = port["Name"]?.ToString();
                            string pHost = port["HostAddress"]?.ToString();
                            if (pName != null && pHost != null) portDict[pName] = pHost;
                        }
                    }
                    Report(); // Шаг 2: Скачали порты

                    using (var printSearcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, PortName, ShareName, Location, Comment FROM Win32_Printer")))
                    {
                        foreach (ManagementObject printer in printSearcher.Get())
                        {
                            string pName = printer["Name"]?.ToString() ?? "";
                            string share = printer["ShareName"]?.ToString() ?? "";
                            string loc = printer["Location"]?.ToString() ?? "";
                            string cmt = printer["Comment"]?.ToString() ?? "";

                            // Проверяем совпадения в памяти программы (моментально)
                            if (pName == invNumber || pName.Contains(invNumber) || pName.Contains(paddedInv) ||
                                share.Contains(paddedInv) || share == invNumber ||
                                loc.Contains(invNumber) || cmt.Contains(invNumber))
                            {
                                string portName = printer["PortName"]?.ToString() ?? "";
                                string ip = portDict.ContainsKey(portName) ? portDict[portName] : portName;

                                progress?.Report(100);
                                return new RemotePrinterInfo { ServerName = server, PrinterName = pName, IpAddress = ip };
                            }
                        }
                    }
                    Report(); // Шаг 3: Проверили принтеры
                }
                catch { Report(); Report(); Report(); /* Если сервер недоступен, проматываем шаги */ }

                Report(); // Шаг 4: Сервер обработан
            }

            progress?.Report(100);
            return null;
        }

        // === СУПЕРБЫСТРЫЙ ПОИСК ПО IP ===
        // === УЛУЧШЕННЫЙ ПОИСК ПО IP ===
        public IpCheckResult CheckIpOnServers(string ipAddress, IProgress<int> progress = null)
        {
            var servers = ConfigService.PrintServers;
            int totalServers = servers.Count;
            if (totalServers == 0) return new IpCheckResult { IpExists = false };

            ipAddress = ipAddress.Trim(); // Убираем пробелы из входящего IP
            int step = 0;
            int totalSteps = totalServers * 2; // Упростим шаги для прогресс-бара

            progress?.Report(0);

            // Список для хранения найденных "пустых" портов (на случай если принтер не найдется)
            IpCheckResult emptyPortResult = null;

            foreach (var server in servers)
            {
                try
                {
                    ConnectionOptions options = new ConnectionOptions
                    {
                        Timeout = TimeSpan.FromSeconds(3),
                        EnablePrivileges = true
                    };
                    ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2", options);
                    scope.Connect();

                    // 1. Кэшируем порты сервера
                    var portDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, HostAddress FROM Win32_TCPIPPrinterPort")))
                    {
                        foreach (ManagementObject port in searcher.Get())
                        {
                            string pName = port["Name"]?.ToString()?.Trim();
                            string pHost = port["HostAddress"]?.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(pName))
                                portDict[pName] = pHost ?? "";
                        }
                    }

                    // 2. Ищем принтеры, привязанные к этому IP
                    using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, PortName FROM Win32_Printer")))
                    {
                        foreach (ManagementObject printer in searcher.Get())
                        {
                            string pName = printer["Name"]?.ToString()?.Trim() ?? "";
                            string portName = printer["PortName"]?.ToString()?.Trim() ?? "";

                            // Проверяем: совпадает ли имя порта с IP или HostAddress порта с IP
                            bool isIpMatch = portName.Equals(ipAddress, StringComparison.OrdinalIgnoreCase) ||
                                            portName.Equals("IP_" + ipAddress, StringComparison.OrdinalIgnoreCase) ||
                                            (portDict.ContainsKey(portName) && portDict[portName].Equals(ipAddress, StringComparison.OrdinalIgnoreCase));

                            if (isIpMatch)
                            {
                                progress?.Report(100);
                                // Нашли реальный принтер - это приоритет, возвращаем сразу
                                return new IpCheckResult
                                {
                                    IpExists = true,
                                    PortName = portName,
                                    ServerName = server,
                                    AttachedPrinterName = pName
                                };
                            }
                        }
                    }

                    // 3. Если принтер не найден, проверяем, нет ли просто "пустого" порта на этом сервере
                    if (emptyPortResult == null)
                    {
                        foreach (var kvp in portDict)
                        {
                            if (kvp.Value.Equals(ipAddress, StringComparison.OrdinalIgnoreCase) ||
                                kvp.Key.Equals(ipAddress, StringComparison.OrdinalIgnoreCase) ||
                                kvp.Key.Equals("IP_" + ipAddress, StringComparison.OrdinalIgnoreCase))
                            {
                                // Запоминаем, но не возвращаем сразу (вдруг на следующем сервере есть живой принтер)
                                emptyPortResult = new IpCheckResult
                                {
                                    IpExists = true,
                                    PortName = kvp.Key,
                                    ServerName = server,
                                    AttachedPrinterName = null
                                };
                            }
                        }
                    }
                }
                catch { /* Игнорируем ошибки подключения к одному серверу */ }

                step += 2;
                progress?.Report(Math.Min((step * 100) / totalSteps, 99));
            }

            progress?.Report(100);

            // Если нашли хоть один пустой порт (и ни одного принтера), возвращаем его
            return emptyPortResult ?? new IpCheckResult { IpExists = false };
        }

        public bool DeletePrinterFromServer(string server, string printerName)
        {
            try
            {
                ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2");
                scope.Connect();
                using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery($"SELECT * FROM Win32_Printer WHERE Name = '{printerName}'")))
                {
                    foreach (ManagementObject printer in searcher.Get())
                    {
                        printer.Delete();
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        public List<string> GetAvailableDrivers(IProgress<int> progress = null)
        {
            HashSet<string> drivers = new HashSet<string>();
            var servers = ConfigService.PrintServers;
            int total = servers.Count;
            if (total == 0) return new List<string>();

            int step = 0;
            int totalSteps = total * 3;
            void Report() { step++; progress?.Report((step * 100) / totalSteps); }

            progress?.Report(0);

            foreach (int i in System.Linq.Enumerable.Range(0, total))
            {
                string server = servers[i];
                try
                {
                    ConnectionOptions options = new ConnectionOptions { Timeout = TimeSpan.FromSeconds(3) };
                    ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2", options);
                    scope.Connect();
                    Report();

                    using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name FROM Win32_PrinterDriver")))
                    {
                        foreach (ManagementObject driver in searcher.Get())
                        {
                            string driverName = driver["Name"]?.ToString();
                            if (!string.IsNullOrEmpty(driverName))
                            {
                                int commaIndex = driverName.IndexOf(',');
                                drivers.Add(commaIndex > 0 ? driverName.Substring(0, commaIndex) : driverName);
                            }
                        }
                    }
                    Report();
                }
                catch { Report(); Report(); }
                Report();
            }

            var list = new List<string>(drivers);
            list.Sort();
            progress?.Report(100);
            return list;
        }

        public bool CreatePrinterOnServer(string server, string printerName, string shareName, string ipAddress, string driverName, string location, IProgress<int> progress = null)
        {
            try
            {
                progress?.Report(10);
                ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2");
                scope.Connect();

                progress?.Report(40);

                // ВОТ ТУТ УБРАЛИ IP_
                string portName = ipAddress;

                ManagementClass portClass = new ManagementClass(scope, new ManagementPath("Win32_TCPIPPrinterPort"), null);
                ManagementObject port = portClass.CreateInstance();
                port["Name"] = portName;
                port["HostAddress"] = ipAddress;
                port["Protocol"] = 1;
                port["PortNumber"] = 9100;
                port.Put();

                progress?.Report(70);
                ManagementClass printerClass = new ManagementClass(scope, new ManagementPath("Win32_Printer"), null);
                ManagementObject printer = printerClass.CreateInstance();
                printer["DeviceID"] = printerName;
                printer["DriverName"] = driverName;
                printer["PortName"] = portName; // Используем то же имя без префикса
                printer["Location"] = location;
                printer["Network"] = true;
                printer["Shared"] = true;
                printer["ShareName"] = shareName;
                printer.Put();

                progress?.Report(100);
                return true;
            }
            catch { progress?.Report(100); return false; }
        }
    }
}