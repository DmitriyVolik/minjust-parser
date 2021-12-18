using System;

namespace minjust_parser.Models
{
    [Serializable]
    public class Config
    {
        public long PersonCounter { get; set; }

        public string FilePathInput { get; set; }

        public string FilePathOutput { get; set; }

        public string CaptchaApiId { get; set; }

        public int ThreadCount { get; set; }
    }
}