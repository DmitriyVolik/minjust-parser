﻿using minjust_parser.Core.Services;
using minjust_parser.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using System.Web;
using DocumentFormat.OpenXml.Spreadsheet;

namespace minjust_parser.Core.Workers
{
    public class ThreadWorker
    {
        public ThreadWorker(ref Config config, ref Captcha captcha, ref List<string> IdentityNumbers, ref List<string> ParsedNumbers)
        {
            this.config = config;
            this.captcha = captcha;
            this.IdentityNumbers = IdentityNumbers;
            this.ParsedNumbers = ParsedNumbers;
        }
        Captcha captcha { get; set; } = null;
        Config config { get; set; } = null;
        List<string> IdentityNumbers { get; set; } = null;
        List<string> ParsedNumbers { get; set; } = null;
        string WorkIdentityNumber { get; set; } = "";
        WebProxy proxy { get; set; } = null;
        Random rand = new Random();
        WebRequest request = null;
        public void StartThread()
        {
            while (IdentityNumbers.Count > 0)
            {
                if (config.IsProxy)
                {
                    if (config.Proxy.Count != 0)
                    {
                        try
                        {
                            Console.WriteLine($"Соединение со случайным прокси сервером...", ConsoleColor.Yellow);
                            proxy = new WebProxy(config.Proxy[rand.Next(0, config.Proxy.Count - 1)]);
                            request = WebRequest.Create("https://google.com");
                            request.Proxy = proxy;
                            WebResponse response1 = request.GetResponse();
                        }
                        catch (Exception)
                        {
                            Console.WriteLine($"Соединение с прокси сервером {proxy.Address} не удалось!", ConsoleColor.Red);
                            continue;
                        }
                        Console.WriteLine($"Соединение с прокси сервером {proxy.Address} установлено!", ConsoleColor.Green);
                    }
                    else
                    {
                        Console.WriteLine("Ошибка. При включенной опции IsProxy количество прокси серверов в списке Proxy не может быть равно нулю!", ConsoleColor.Red);
                        throw new Exception();
                    }
                }

                WorkIdentityNumber = IdentityNumbers[0];
                var temp = IdentityNumbers[0];
                IdentityNumbers.RemoveAt(0);

                Console.WriteLine($"Осталось парсить {IdentityNumbers.Count} номеров", ConsoleColor.Gray);

          


                try
                {
                    if (!ParsedNumbers.Contains(WorkIdentityNumber))
                    {

                        Console.WriteLine($"Решаю капчу для {WorkIdentityNumber}...");

                  
                        

                        var captchaToken = captcha.SolveReCaptcha();

                        Console.WriteLine($"Капча для {WorkIdentityNumber} решена!", ConsoleColor.Green);
                        Console.WriteLine($"Начинаю парсинг данных для {WorkIdentityNumber}.");


                        request = WebRequest.Create($"https://usr.minjust.gov.ua/USRWebAPI/api/public/search?person={WorkIdentityNumber}&c={captchaToken}");

                        request.Proxy = proxy;

                        if (config.IsProxy) request.Proxy = proxy;

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
                            request = WebRequest.Create($"https://usr.minjust.gov.ua/USRWebAPI/api/public/detail?rfId={HttpUtility.UrlEncode(output[i])}");
                            response = request.GetResponse();

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
                            if (person.Count != 0)
                            {
                                try
                                {
                                    Excel.Write(person, config.FilePathOutput, config.PersonOutCounter + 2, WorkIdentityNumber);
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Файл для записи не найден, укажите правильный путь в config.json", ConsoleColor.Red);
                                    throw;
                                }

                                Console.WriteLine($"Данные { WorkIdentityNumber} занесены в {config.FilePathOutput} файл.", ConsoleColor.Green);
                                config.PersonOutCounter++;
                                ParsedNumbers.Add(WorkIdentityNumber);
                                FileWorker.WriteParsedNumber(WorkIdentityNumber);
                                FileWorker.SaveConfig(config);
                            }
                        }
                        if (output.Count==0)
                        {
                            ParsedNumbers.Add(WorkIdentityNumber);
                            FileWorker.WriteParsedNumber(WorkIdentityNumber);
                        }
                    }
                    else
                    {
                        Console.WriteLine($"{WorkIdentityNumber} уже содержится в {config.FilePathOutput}");
                        continue;
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine($"Ошибка подключения. {temp} добавлен в конец очереди. Парсинг другого элемента");
                    IdentityNumbers.Add(temp);
                    continue;
                }
            }
            Console.WriteLine($"Поток {Thread.CurrentThread.ManagedThreadId} завершил работу.");
        }
    }
}