﻿using minjust_parser.Core.Services;
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
            IdNumbers=Excel.Read();
        }
        public Config config = null;
        public Captcha captcha = null;

        public List<string> IdNumbers = new List<string>();

        public void Start()
        {
            Excel.WriteStartPattern(config.FilePathOutput);

            for (int i = 0; i < 1; i++)
            {
                ThreadWorker worker = new ThreadWorker(ref config, ref captcha, ref IdNumbers);
                Thread thread = new Thread(new ThreadStart(worker.StartThread));
                thread.Start();
                Thread.Sleep(1000);
            }
        }
    }
}
