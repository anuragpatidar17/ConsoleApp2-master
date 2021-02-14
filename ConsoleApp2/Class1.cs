using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Reflection;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace ConsoleApp2
{

    public class Class1
    {
        [Test]
        public void SearchForWord()
        {
            var driver = new ChromeDriver();

            {
           
            //FileStream file = new FileStream(@"/Users/anuragpatidar/Desktop/ConsoleApp2-master/Book1.xlsx", FileMode.Open, FileAccess.Read);
                FileStream file = new FileStream(@"D:\a\1\s\Book1.xlsx", FileMode.Open, FileAccess.Read);
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheet("Sheet1");

                var value = string.Format(sheet.GetRow(0).GetCell(0).StringCellValue);



                driver.Navigate().GoToUrl("http://www.google.com/");


                IWebElement query = driver.FindElement(By.Name("q"));


                Console.WriteLine(driver.Title + value);

                query.SendKeys(value);


                query.Submit();



                driver.Quit();
            }

        }
    }
}
