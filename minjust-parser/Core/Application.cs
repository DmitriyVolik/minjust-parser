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
            Console.WriteLine(captchaToken);
            
            WebRequest request = WebRequest.Create($"https://usr.minjust.gov.ua/USRWebAPI/api/public/search?person=0000000001&c={captchaToken}");
            WebResponse response = request.GetResponse();

            string responseStr;
            
            using (Stream stream = response.GetResponseStream())
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    responseStr = reader.ReadToEnd();
                }     
            }
            response.Close();

            var output = Helpers.GetRFID(responseStr);
            
            for (int i = 0; i < output.Count; i++)
            {
                request=WebRequest.Create($"https://usr.minjust.gov.ua/USRWebAPI/api/public/detail?rfId={HttpUtility.UrlEncode(output[i])}");
                response=request.GetResponse();
                
                string tempResponse;
                
                using (Stream stream = response.GetResponseStream())
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        tempResponse = reader.ReadToEnd();
                    }     
                }
                response.Close();
                
                Console.WriteLine(tempResponse);

                var person = JsonWorker<List<PersonData>>.JsonToObj(tempResponse);

                foreach (var item in person)
                {
                    Console.WriteLine(item.value);
                }
                
            }

            Console.WriteLine("Запрос выполнен");

            ThreadWorker worker = new ThreadWorker(ParserSettings.TestValue);
            Thread thread = new Thread(new ThreadStart(worker.StartThread));
            thread.Start();

        }
    }
}
