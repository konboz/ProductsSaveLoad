using System;

namespace SaveLoad
{
    class Program
    {
        static void Main(string[] args)
        {
            var product1 = new Product(1, "Shirt", 29.99m, ProductCategory.Clothing);
            var product2 = new Product(2, "Beer", 3.65m, ProductCategory.Drinks );
            var basket1 = new Basket();
            basket1.groceries.Add(product1);
            basket1.groceries.Add(product2);

            ISerializer serializer = new ExcelSerializer();
            serializer.SaveToFile($"{nameof(basket1)}", basket1);

            var basketFromFile = serializer.ReadFromFile($"{nameof(basket1)}");
           
            Console.ReadLine();
        }
    }
}
