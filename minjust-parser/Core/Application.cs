using minjust_parser.Core.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using minjust_parser.Core.Workers;
using minjust_parser.Models;


namespace minjust_parser.Core
{
    public class Application
    {

        public Application()
        {
            config = FileWorker.LoadConfig().Result;
            
            try
            {
                captcha = new Captcha(config.CaptchaApiId);
            }
            catch (Exception)
            {
                throw;
            }
            try
            {
                IdNumbers=Excel.Read(config);
            }
            catch (Exception e)
            {
                Console.WriteLine("Входной файл не найден, укажите правильный путь в config.json");
                throw;
            }
        }
        public Config config = null;
        public Captcha captcha = null;

        public List<string> IdNumbers = new List<string>();
        public List<string> ParsedNumbers = new List<string>();
        public void Start()
        {
            Excel.WriteStartPattern(config.FilePathOutput);

            for (int i = 0; i < config.ThreadCount; i++)
            {
                ThreadWorker worker = new ThreadWorker(ref config, ref captcha, ref IdNumbers, ref ParsedNumbers);
                Thread thread = new Thread(new ThreadStart(worker.StartThread));
                thread.Start();
                Thread.Sleep(1000);
            }
        }
    }
}
