using System.Collections.Generic;

namespace minjust_parser.Core.Services
{
    public class Helpers
    {
        public static List<string> GetInfoUrl(string input, string data)
        {
            var array = input.Split($"\"{data}\":\"");
            List<string> response = new List<string>();
            for (int i = 1; i < array.Length; i++)
            {
                array[i] = array[i].Remove(array[i].IndexOf('"'), array[i].Length - array[i].IndexOf('"'));
                response.Add(array[i]);
            }
            return response;
        }
    }
}