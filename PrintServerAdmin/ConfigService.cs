using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
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
        public string PrintServer { get; set; }
    }

    public static class ConfigService
    {
        public static List<string> PrintServers { get; private set; } = new List<string>();
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

                if (File.Exists(serversFile))
                {
                    string serversJson = File.ReadAllText(serversFile);
                    PrintServers = JsonConvert.DeserializeObject<List<string>>(serversJson);
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
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка чтения конфигов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}