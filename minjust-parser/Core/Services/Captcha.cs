using TwoCaptcha;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TwoCaptcha.Captcha;

namespace minjust_parser.Core.Services
{
    public class Captcha
    {
        public Captcha(string ApiKey)
        {
            this.ApiKey = ApiKey;
            client = new TwoCaptcha.TwoCaptcha(this.ApiKey);
        }
        public string ApiKey { get; set; }
        public TwoCaptcha.TwoCaptcha client { get; set; } = null;

        public string SolveReCaptcha()
        {
            ReCaptcha captcha = new ReCaptcha();
            captcha.SetSiteKey("6LdStXoUAAAAAE2oEyZLHgu3dBE-WV1zOvZon7_v");
            captcha.SetUrl("https://usr.minjust.gov.ua/content/free-search");
            try
            {
                client.Solve(captcha).Wait();
                return captcha.Code;
            }
            catch (Exception)
            {
                Console.WriteLine("Captcha Error");
                return null;
            }
        }
    }
}
