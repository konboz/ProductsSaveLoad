using System;

namespace SaveLoad
{
    class Program
    {
        static void Main(string[] args)
        {
            Basket basket = new Basket();
            basket.FillWithDummyData();
            Console.WriteLine(basket);
            basket.SaveToText();
            basket.SaveToJson();
            basket.SaveToExcel();
            basket.LoadFromText();
            //basket.LoadFromJson();
            //basket.LoadFromExcel();
            Console.WriteLine(basket);
            Console.ReadLine();
        }
    }
}
