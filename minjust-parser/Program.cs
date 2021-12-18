using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using minjust_parser.Core;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace minjust_parser
{
    class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application();
            app.Start();
        }
    }
}