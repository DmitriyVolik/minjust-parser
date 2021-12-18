using minjust_parser.Core.Services;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using TwoCaptcha.Captcha;
using Captcha = minjust_parser.Core.Services.Captcha;

namespace minjust_parser.Core
{
    public class Application
    {
        public Application()
        {
            captcha = new Captcha(ParserSettings.captchaServiceKey);
            IdNumbers=Excel.Read();
        }
        public Captcha captcha { get; set; }

        public List<string> IdNumbers { get; set; }

        public void Start()
        {
            var captchaToken = captcha.SolveReCaptcha();
            //Console.WriteLine(captchaToken);
            
            WebRequest request = WebRequest.Create($"https://usr.minjust.gov.ua/USRWebAPI/api/public/search?person={IdNumbers[0]}&c={captchaToken}");
            WebResponse response = request.GetResponse();
            
            string responseStr = "";
            
            using (Stream stream = response.GetResponseStream())
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    
                    while ((reader.ReadLine()) != null)
                    {
                        responseStr += reader.ReadLine();
                    }
                }     
            }
            Console.WriteLine(responseStr); 

            Console.WriteLine(HttpUtility.UrlEncode("Y7jDYHydLkYF0Nmb8KKcf6QOqZEFdYSFfjlbUmMl7zZyViv10jHCs8P+hp6pujWkfOpgcFplKcb5mJrYN4uOEA=="));
            //response.Close();
            Console.WriteLine("Запрос выполнен");
        }
    }
}
