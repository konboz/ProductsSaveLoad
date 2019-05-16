using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SaveLoad
{
    class ExcelSerializer : ISerializer
    {
        public Basket ReadFromFile(string fileName)
        {
            var basket = new Basket();
            var groceries = new List<Product>();
            try
            {
                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream($@"{fileName}.xlsx", FileMode.Open,
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
                        decimal price = decimal.Parse(sheet.GetRow(row).GetCell(2).ToString());
                        var category = Enum.Parse<ProductCategory>(sheet.GetRow(row).GetCell(3).StringCellValue);

                        Product a = new Product(id, name, price, category);
                        groceries.Add(a);
                    }

                }
                basket.groceries = groceries;

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return basket;
        }

        public bool SaveToFile(string fileName, Basket basket)
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Mysheet");

            var r = sheet.CreateRow(0);

            r.CreateCell(0).SetCellValue("Id No");
            r.CreateCell(1).SetCellValue("Name");
            r.CreateCell(2).SetCellValue("Price");
            r.CreateCell(3).SetCellValue("Category");
            for (int i = 0; i < basket.groceries.Count; i++)
            {
                r = sheet.CreateRow(i + 1);
                r.CreateCell(0).SetCellValue(basket.groceries[i].Id);
                r.CreateCell(1).SetCellValue(basket.groceries[i].Name);
                r.CreateCell(2).SetCellValue(basket.groceries[i].Price.ToString());
                r.CreateCell(3).SetCellValue(basket.groceries[i].Category.ToString());
            }
            try
            {
                using (var fs = new FileStream($@"{fileName}.xlsx", FileMode.Create,
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
    }
}
