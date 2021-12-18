using System.Text.Json;

namespace minjust_parser.Core.Services
{
    public static class JsonWorker<T>
    {
        public static string ObjToJson(T obj)
        {
            var settings = new JsonSerializerOptions()
            {
                WriteIndented = true
            };
            return JsonSerializer.Serialize(obj, settings);
        }
        public static T JsonToObj(string jsonData)
        {
            return JsonSerializer.Deserialize<T>(jsonData);
        }
    
    }
}