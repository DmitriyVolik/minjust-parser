using System;
using System.Collections.Generic;

namespace minjust_parser.Models
{
    [Serializable]
    public class Config
    {
        public long PersonOutCounter { get; set; } = 0;

        public string FilePathInput { get; set; } = "input.xlsx";

        public string FilePathOutput { get; set; } = "output.xlsx";

        public string CaptchaApiId { get; set; } = "";

        public int ThreadCount { get; set; } = 1;
        public bool IsProxy { get; set; } = false;
        public List<string> Proxy { get; set; } = new List<string>()
        {
            ""
        };
    }
}