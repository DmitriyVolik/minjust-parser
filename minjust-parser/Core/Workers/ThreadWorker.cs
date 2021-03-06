using minjust_parser.Core.Services;
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
        public ThreadWorker(ref Config config, ref Captcha captcha, ref List<string> Names, ref List<string> ParsedNames)
        {
            this.config = config;
            this.captcha = captcha;
            this.Names = Names;
            this.ParsedNames = ParsedNames;
        }
        Captcha captcha { get; set; } = null;
        Config config { get; set; } = null;
        List<string> Names { get; set; } = null;
        List<string> ParsedNames { get; set; } = null;
        string SearchName { get; set; } = "";
        WebProxy proxy { get; set; } = null;
        Random rand = new Random();
        WebRequest request = null;
        public void StartThread()
        {
            while (Names.Count > 0)
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

                SearchName = Names[0];
                var temp = Names[0];
                Names.RemoveAt(0);

                Console.WriteLine($"Осталось парсить {Names.Count} значений", ConsoleColor.Gray);

                try
                {
                    if (!ParsedNames.Contains(SearchName))
                    {
                        Console.WriteLine($"Решаю капчу для {SearchName}...");

                        var captchaToken = captcha.SolveReCaptcha();
                        
                        Console.WriteLine($"Капча для {SearchName} решена!", ConsoleColor.Green);
                        Console.WriteLine($"Начинаю парсинг данных для {SearchName}.");

                        request = WebRequest.Create($"https://usr.minjust.gov.ua/USRWebAPI/api/public/search?person={HttpUtility.UrlEncode(SearchName)}&c={captchaToken}");

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

                        var outputRfIds = Helpers.GetInfoUrl(responseStr, "rfId");
                        var outputStates = Helpers.GetInfoUrl(responseStr, "state");

                        for (int i = 0; i < outputRfIds.Count; i++)
                        {
                            
                            if (outputStates[i]=="припинено")
                            {
                                continue;
                            }
                            
                            request = WebRequest.Create($"https://usr.minjust.gov.ua/USRWebAPI/api/public/detail?rfId={HttpUtility.UrlEncode(outputRfIds[i])}");
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
                            var person = JsonWorker<List<PersonData>>.JsonToObj(tempResponse);
                            if (person.Count != 0)
                            {
                                try
                                {
                                    Excel.Write(person, config.FilePathOutput, config.PersonOutCounter + 2, SearchName);
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Файл для записи не найден, укажите правильный путь в config.json", ConsoleColor.Red);
                                    throw;
                                }

                                Console.WriteLine($"Данные { SearchName} занесены в {config.FilePathOutput} файл.", ConsoleColor.Green);

                                Console.WriteLine($"{SearchName}:ЗАПИСАН", ConsoleColor.Yellow);

                                config.PersonOutCounter++;
                                
                                FileWorker.SaveConfig(config);
                            }
                        }
                        if (outputRfIds.Count==0)
                        {
                            Console.WriteLine($"{SearchName}: ИДЕНТИФИКАЦИОННЫЙ НОМЕР НЕ СОДЕРЖИТ ИНФОРМАЦИИ (ПУСТ)", ConsoleColor.Red);
                        }
                        ParsedNames.Add(SearchName);
                        FileWorker.WriteParsedName(SearchName);
                    }
                    else
                    {
                        Console.WriteLine($"{SearchName} уже содержится в {config.FilePathOutput}");
                        continue;
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine($"Ошибка подключения. {temp} добавлен в конец очереди. Парсинг другого элемента");
                    Names.Add(temp);
                    continue;
                }
            }
            
            Console.WriteLine($"Поток {Thread.CurrentThread.ManagedThreadId} завершил работу.");
        }
    }
}