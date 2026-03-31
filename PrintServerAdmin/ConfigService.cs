using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;

namespace PrintServerAdmin
{
    public class City
    {
        public string Name { get; set; }
        public string Code { get; set; }
    }

    public class ServerMapping
    {
        public string CityCode { get; set; } // Идеально совпадает с твоим файлом
        public string PrinterType { get; set; }
        public string PrintServer { get; set; }
    }

    public static class ConfigService
    {
        public static List<string> PrintServers { get; private set; } = new List<string>();
        public static List<string> TpPrintServers { get; private set; } = new List<string>();
        public static List<City> Cities { get; private set; } = new List<City>();
        public static List<ServerMapping> Mappings { get; private set; } = new List<ServerMapping>();

        public static void LoadConfigs()
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string serversFile = Path.Combine(baseDir, "servers_list.json");
                string citiesFile = Path.Combine(baseDir, "cities.json");
                string mapFile = Path.Combine(baseDir, "server_mapping.json");
                string legacyMapFile = Path.Combine(baseDir, "servers_mapping.json");

                if (File.Exists(serversFile))
                {
                    string serversJson = File.ReadAllText(serversFile);
                    ParseServersConfig(serversJson);
                    if (PrintServers.Count == 0) PrintServers.Add("LG166PS");
                }
                else PrintServers.Add("LG166PS");

                if (File.Exists(citiesFile))
                {
                    string citiesJson = File.ReadAllText(citiesFile);
                    Cities = JsonConvert.DeserializeObject<List<City>>(citiesJson);
                }
                else MessageBox.Show($"Файл cities.json не найден в папке {baseDir}", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                if (File.Exists(mapFile))
                {
                    string mappingJson = File.ReadAllText(mapFile);
                    Mappings = JsonConvert.DeserializeObject<List<ServerMapping>>(mappingJson);
                }
                else if (File.Exists(legacyMapFile))
                {
                    string mappingJson = File.ReadAllText(legacyMapFile);
                    Mappings = JsonConvert.DeserializeObject<List<ServerMapping>>(mappingJson);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка чтения конфигов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ParseServersConfig(string serversJson)
        {
            var token = JToken.Parse(serversJson);
            PrintServers = new List<string>();
            TpPrintServers = new List<string>();

            // Backward compatibility: old format ["srv1", "srv2"]
            if (token.Type == JTokenType.Array)
            {
                PrintServers = token.ToObject<List<string>>() ?? new List<string>();
                return;
            }

            if (token.Type == JTokenType.Object)
            {
                JObject root = (JObject)token;
                PrintServers = root["default"]?.ToObject<List<string>>() ?? new List<string>();
                TpPrintServers = root["tp"]?.ToObject<List<string>>() ?? new List<string>();

                if (PrintServers.Count == 0 && root["servers"] != null)
                    PrintServers = root["servers"].ToObject<List<string>>() ?? new List<string>();

                return;
            }
        }
    }
}