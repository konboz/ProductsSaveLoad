using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SaveLoad
{
    class JsonSerializer : ISerializer
    {
        public Basket ReadFromFile(string fileName)
        {
            var groceries = new List<Product>();
            var basket = new Basket();
            try
            {
                string data = File.ReadAllText($@"{fileName}.json");
                basket = JsonConvert.DeserializeObject<Basket>(data);
            }
            catch (Exception e)
            {
                return null;
            }
            return basket;
        }


        public bool SaveToFile(string fileName, Basket basket)
        {
            try
            {
                var json = JsonConvert.SerializeObject(basket, Formatting.None);
                //write string to file
                File.WriteAllText($@"{ fileName}.json", json);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }
    }
}
