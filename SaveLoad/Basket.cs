using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SaveLoad
{
    public class Basket
    {
        public List<Product> Groceries = new List<Product>();

        private const string filenametxt = @"Groceries.txt";
        private const string filenamejson = @"Groceries.json";
        private const string filenamexlsx = @"Groceries.xlsx";

        public void FillWithDummyData()
        {
            Groceries.Add(new Product(1, "Shirt", 29.99, "Clothing"));
            Groceries.Add(new Product(2, "Apples", 2.45, "Fruits"));
            Groceries.Add(new Product(3, "Brocolli", 4.65, "Vegetables"));
            Groceries.Add(new Product(4, "Batteries", 3.92, "Electronics"));
            Groceries.Add(new Product(5, "Shampoo", 4.25, "Toiletries"));
            Groceries.Add(new Product(6, "Beer", 3.65, "Drinks"));
        }

        public override string ToString()
        {
            foreach (Product t in Groceries)
            {
                Console.WriteLine(t);
            }
            return "";
        }

        public void SaveToText() //Works with privet Product properties
        {

            using (StreamWriter file = new StreamWriter(filenametxt))
            {
                foreach (Product t in Groceries)
                {
                    file.WriteLine(t);
                }
            }
        }

        public void SaveToJson() //Doesn't work with private Product properties
        {
            string json = JsonConvert.SerializeObject(Groceries.ToArray());
            //write string to file
            File.WriteAllText(filenamejson, json);
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
            for (int i = 0; i < Groceries.Count; i++)
            {
                r = sheet.CreateRow(i + 1);
                r.CreateCell(0).SetCellValue(Groceries[i].Id);
                r.CreateCell(1).SetCellValue(Groceries[i].Name);
                r.CreateCell(2).SetCellValue(Groceries[i].Price);
                r.CreateCell(3).SetCellValue(Groceries[i].Category);
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
            catch (Exception)
            {
                return false;
            }
        }

        
        public bool LoadFromText() //Doesn't work at this stage
        {
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
                            Console.WriteLine($"{item[0]}, {item[1]}, {item[2]}, {item[3]}");
                            Product a = new Product(int.Parse(item[0]), item[1], double.Parse(item[2]), item[3]);
                            Console.WriteLine(a);
                            //Groceries.Add(new Product(int.Parse(item[0]), item[1], double.Parse(item[2]), item[3]));
                        }
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void LoadFromJson()
        {
            string data = File.ReadAllText(filenamejson);
            var tempGroceries = JsonConvert.DeserializeObject<List<Product>>(data);
            foreach (Product t in tempGroceries)
            {

                Groceries.Add(t);
            }
        }

        public bool LoadFromExcel()
        {
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
                        Groceries.Add(a);
                    }

                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
