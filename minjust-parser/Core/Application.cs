using minjust_parser.Core.Services;
using minjust_parser.Core.Workers;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace minjust_parser.Core
{
    public class Application
    {
        public Captcha captcha { get; set; }
        public void Start()
        {
            ThreadWorker worker = new ThreadWorker(ParserSettings.TestValue);
            Thread thread = new Thread(new ThreadStart(worker.StartThread));
            thread.Start();
        }
    }
}
