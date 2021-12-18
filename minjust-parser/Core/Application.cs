using minjust_parser.Core.Services;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace minjust_parser.Core
{
    public class Application
    {
        public Application()
        {
            captcha = new Captcha(ParserSettings.captchaServiceKey);
        }
        public Captcha captcha { get; set; }
        public void Start()
        {
            var token = captcha.SolveReCaptcha();
            Console.WriteLine(token);
        }
    }
}
