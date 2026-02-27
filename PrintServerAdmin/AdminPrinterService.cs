using System;
using System.Management;
using System.Collections.Generic;
using System.Windows.Forms; // Добавили для вывода ошибок на экран

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
        // ================= ПОИСК ПО ИНВЕНТАРНИКУ (С ОТЛАДКОЙ) =================
        public RemotePrinterInfo FindPrinterForDeletion(string invNumber)
        {
            string paddedInv = invNumber.PadLeft(6, '0');
            string errorLog = "";

            foreach (var server in ConfigService.PrintServers)
            {
                try
                {
                    ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2");
                    scope.Connect(); // Если тут нет прав, мы об этом узнаем

                    using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, PortName, ShareName, Location, Comment FROM Win32_Printer")))
                    {
                        foreach (ManagementObject printer in searcher.Get())
                        {
                            string pName = printer["Name"]?.ToString() ?? "";
                            string share = printer["ShareName"]?.ToString() ?? "";
                            string loc = printer["Location"]?.ToString() ?? "";
                            string cmt = printer["Comment"]?.ToString() ?? "";

                            if (pName.Contains(paddedInv) || share.Contains(paddedInv) ||
                                pName.Contains(invNumber + "-") || share == invNumber ||
                                loc.Contains(invNumber) || cmt.Contains(invNumber))
                            {
                                string portName = printer["PortName"]?.ToString();
                                string ip = GetIpFromPort(scope, portName) ?? portName;
                                return new RemotePrinterInfo { ServerName = server, PrinterName = pName, IpAddress = ip };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorLog += $"Сервер {server}: {ex.Message}\n";
                }
            }

            if (!string.IsNullOrEmpty(errorLog))
            {
                MessageBox.Show("Скрытые ошибки доступа WMI при поиске инв. номера:\n\n" + errorLog, "Отладка WMI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return null;
        }

        // ================= ПОИСК ПО IP (С ОТЛАДКОЙ) =================
        public IpCheckResult CheckIpOnServers(string ipAddress)
        {
            string errorLog = "";

            foreach (var server in ConfigService.PrintServers)
            {
                try
                {
                    ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2");
                    scope.Connect();

                    // 1. Ищем прямо в именах портов принтеров (самый надежный способ)
                    using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, PortName FROM Win32_Printer")))
                    {
                        foreach (ManagementObject printer in searcher.Get())
                        {
                            string portName = printer["PortName"]?.ToString() ?? "";
                            if (portName.Contains(ipAddress))
                            {
                                return new IpCheckResult
                                {
                                    IpExists = true,
                                    PortName = portName,
                                    ServerName = server,
                                    AttachedPrinterName = printer["Name"]?.ToString()
                                };
                            }
                        }
                    }

                    // 2. Ищем в физических TCP/IP портах
                    using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery("SELECT Name, HostAddress FROM Win32_TCPIPPrinterPort")))
                    {
                        foreach (ManagementObject port in searcher.Get())
                        {
                            string pName = port["Name"]?.ToString() ?? "";
                            string host = port["HostAddress"]?.ToString() ?? "";

                            if (host == ipAddress || pName.Contains(ipAddress))
                            {
                                // Порт найден, ищем принтер на этом порту
                                string attachedPrinter = null;
                                using (var pSearcher = new ManagementObjectSearcher(scope, new ObjectQuery($"SELECT Name FROM Win32_Printer WHERE PortName = '{pName}'")))
                                {
                                    foreach (ManagementObject p in pSearcher.Get())
                                    {
                                        attachedPrinter = p["Name"]?.ToString();
                                        break;
                                    }
                                }
                                return new IpCheckResult { IpExists = true, PortName = pName, ServerName = server, AttachedPrinterName = attachedPrinter };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    errorLog += $"Сервер {server}: {ex.Message}\n";
                }
            }

            if (!string.IsNullOrEmpty(errorLog))
            {
                MessageBox.Show("Скрытые ошибки доступа WMI при проверке IP:\n\n" + errorLog, "Отладка WMI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return new IpCheckResult { IpExists = false };
        }

        public bool DeletePrinterFromServer(string server, string printerName)
        {
            try
            {
                ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2");
                scope.Connect();
                string query = $"SELECT * FROM Win32_Printer WHERE Name = '{printerName}'";
                using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery(query)))
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

        public List<string> GetAvailableDrivers()
        {
            HashSet<string> drivers = new HashSet<string>();
            foreach (var server in ConfigService.PrintServers)
            {
                try
                {
                    ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2");
                    scope.Connect();
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
                }
                catch { }
            }
            var list = new List<string>(drivers);
            list.Sort();
            return list;
        }

        public bool CreatePrinterOnServer(string server, string printerName, string shareName, string ipAddress, string driverName, string location)
        {
            try
            {
                ManagementScope scope = new ManagementScope($@"\\{server}\root\cimv2");
                scope.Connect();

                string portName = "IP_" + ipAddress;
                ManagementClass portClass = new ManagementClass(scope, new ManagementPath("Win32_TCPIPPrinterPort"), null);
                ManagementObject port = portClass.CreateInstance();
                port["Name"] = portName;
                port["HostAddress"] = ipAddress;
                port["Protocol"] = 1;
                port["PortNumber"] = 9100;
                port.Put();

                ManagementClass printerClass = new ManagementClass(scope, new ManagementPath("Win32_Printer"), null);
                ManagementObject printer = printerClass.CreateInstance();
                printer["DeviceID"] = printerName;
                printer["DriverName"] = driverName;
                printer["PortName"] = portName;
                printer["Location"] = location;
                printer["Network"] = true;
                printer["Shared"] = true;
                printer["ShareName"] = shareName;
                printer.Put();

                return true;
            }
            catch { return false; }
        }

        private string GetIpFromPort(ManagementScope scope, string portName)
        {
            try
            {
                string query = $"SELECT HostAddress FROM Win32_TCPIPPrinterPort WHERE Name = '{portName}'";
                using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery(query)))
                {
                    foreach (ManagementObject port in searcher.Get())
                        return port["HostAddress"]?.ToString();
                }
                return portName;
            }
            catch { }
            return null;
        }
    }
}