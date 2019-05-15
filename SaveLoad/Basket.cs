using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace SaveLoad
{
    public class Basket
    {
        public List<Product> groceries = new List<Product>();

        private const string filenametxt = @"Groceries.txt";
        private const string filenamejson = @"Groceries.json";
        private const string filenamexlsx = @"Groceries.xlsx";

        public void FillWithDummyData()
        {
            groceries.Add(new Product(1, "Shirt", 29.99, "Clothing"));
            groceries.Add(new Product(2, "Apples", 2.45, "Fruits"));
            groceries.Add(new Product(3, "Brocolli", 4.65, "Vegetables"));
            groceries.Add(new Product(4, "Batteries", 3.92, "Electronics"));
            groceries.Add(new Product(5, "Shampoo", 4.25, "Toiletries"));
            groceries.Add(new Product(6, "Beer", 3.65, "Drinks"));
        }

        public override string ToString()
        {
            foreach (Product t in groceries)
            {
                Console.WriteLine($"Id: {t.Id}, name: {t.Name}, price: {t.Price}, category: {t.Category}");
            }
            return "";
        }

        public bool SaveToText() //Works with privet Product properties// Saves only the values
        {
            try
            {
                using (StreamWriter file = new StreamWriter(filenametxt))
                {
                    foreach (Product t in groceries)
                    {
                        file.WriteLine(t);
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        public bool SaveToJson() //Doesn't work with private Product properties
        {
            try
            {
                string json = JsonConvert.SerializeObject(groceries.ToArray());
                //write string to file
                File.WriteAllText(filenamejson, json);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        public bool SaveToExcel()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Mysheet");

            var r = sheet.CreateRow(0);

            r.CreateCell(0).SetCellValue("Id No");
            r.CreateCell(1).SetCellValue("Name");
            r.CreateCell(2).SetCellValue("Price");
            r.CreateCell(3).SetCellValue("Category");
            for (int i = 0; i < groceries.Count; i++)
            {
                r = sheet.CreateRow(i + 1);
                r.CreateCell(0).SetCellValue(groceries[i].Id);
                r.CreateCell(1).SetCellValue(groceries[i].Name);
                r.CreateCell(2).SetCellValue(groceries[i].Price);
                r.CreateCell(3).SetCellValue(groceries[i].Category);
            }
            try
            {
                using (var fs = new FileStream(filenamexlsx, FileMode.Create,
                FileAccess.Write))
                {
                    wb.Write(fs);
                }
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }


        public bool LoadFromText()
        {
            groceries.Clear();
            try
            {
                //Read each line from the file and create a Product instance
                foreach (var line in File.ReadAllLines(filenametxt))
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        var item = line.Split(',');
                        if (item.Length == 4)
                        {
                            groceries.Add(new Product(int.Parse(item[0]), item[1], double.Parse(item[2]), item[3]));

                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        public void LoadFromJson()
        {
            groceries.Clear();
            string data = File.ReadAllText(filenamejson);
            var tempGroceries = JsonConvert.DeserializeObject<List<Product>>(data);
            foreach (Product t in tempGroceries)
            {

                groceries.Add(t);
            }
        }

        public bool LoadFromExcel()
        {
            groceries.Clear();
            try
            {
                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(filenamexlsx, FileMode.Open,
                FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }
                ISheet sheet = hssfwb.GetSheet("Mysheet");

                // first row not included, it contains headers
                for (int row = 1; row <= sheet.LastRowNum; row++)
                {
                    if (sheet.GetRow(row) != null)
                    {

                        int id = int.Parse(sheet.GetRow(row).GetCell(0).ToString());
                        string name = sheet.GetRow(row).GetCell(1).ToString();
                        double price = double.Parse(sheet.GetRow(row).GetCell(2).ToString());
                        string category = sheet.GetRow(row).GetCell(3).ToString();

                        Product a = new Product(id, name, price, category);
                        groceries.Add(a);
                    }

                }
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
