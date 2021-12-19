using System;
using System.Collections.Generic;

namespace minjust_parser.Models
{
    [Serializable]
    public class Config
    {
        public long PersonCounter { get; set; } = 2;

        public string FilePathInput { get; set; }

        public string FilePathOutput { get; set; }

        public string CaptchaApiId { get; set; }

        public int ThreadCount { get; set; }
        public List<string> Proxy { get; set; } = new List<string>()
        {
            "188.134.90.77:8080",
            "71.19.250.92:9030",
            "198.8.94.170:4145",
            "85.195.104.71:80",
        };
    }
}