using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
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
        
        
        
        public static List<string> ReadParsedNames()
        {
            List<string> data = new List<string>(); 

            if (!File.Exists("ParsedNames.txt"))
            {
                File.Create("ParsedNames.txt");
                return data;
            }

            string str;

            using (StreamReader sr = new StreamReader("ParsedNames.txt"))
            {
                while ((str=sr.ReadLine())!=null)
                {
                    data.Add(str);
                }
            }

            return data;
        }

        public static void WriteParsedName(string number)
        {
            File.AppendAllTextAsync("ParsedNames.txt", number + '\n');
        }
    }
}