using minjust_parser.Core;
using minjust_parser.Core.Workers;
using System.IO;

namespace minjust_parser
{
    class Program
    {
        static void Main(string[] args)
        {
            FileWorker.CreateAllDirestories();
            Application app = new Application();
            app.Start();
        }
    }
    public static class Console
    {
        public static void WriteLine(dynamic input, System.ConsoleColor color = System.ConsoleColor.White)
        {
            string tempstr = $"[{System.DateTime.Now}]: {input}";
            File.AppendAllTextAsync($"Logs/{System.DateTime.Now.Date.ToShortDateString()}.txt", tempstr + '\n');
            System.Console.ForegroundColor = color;
            System.Console.WriteLine(tempstr);
            System.Console.ResetColor();
        }
    }
}