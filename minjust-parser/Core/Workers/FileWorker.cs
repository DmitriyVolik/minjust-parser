using System.IO;
using System.Threading.Tasks;
using minjust_parser.Core.Services;
using minjust_parser.Models;

namespace minjust_parser.Core.Workers
{
    public static class FileWorker
    {
        public static void CreateAllDirestories()
        {
            Directory.CreateDirectory("Logs");
        }
        public static async Task<Config> LoadConfig()
        {
            if (!File.Exists("config.json"))
            {
                SaveConfig(new Config());
            }
            return JsonWorker<Config>.JsonToObj(await File.ReadAllTextAsync("config.json"));
        }

        public static async void SaveConfig(Config config)
        {
            await File.WriteAllTextAsync("config.json", JsonWorker<Config>.ObjToJson(config));
        }
    }
}