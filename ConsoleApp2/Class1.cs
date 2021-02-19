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
using System.Threading;

namespace ConsoleApp2
{

    public class Class1
    {
        [Test]
        public void SearchForWord()
        {
            var driver = new ChromeDriver();
            driver.Manage().Window.Maximize();

            {
           
            //FileStream file = new FileStream(@"/Users/anuragpatidar/Desktop/ConsoleApp2-master/Book1.xlsx", FileMode.Open, FileAccess.Read);
                //FileStream file = new FileStream(@"D:\a\1\s\Book1.xlsx", FileMode.Open, FileAccess.Read);
                //XSSFWorkbook workbook = new XSSFWorkbook(file);
                //ISheet sheet = workbook.GetSheet("Sheet1");

                //var value = string.Format(sheet.GetRow(0).GetCell(0).StringCellValue);



                driver.Navigate().GoToUrl("http://www.google.com/");


                IWebElement query = driver.FindElement(By.Name("q"));


                string excel_demo = "#rso > div:nth-child(1) > div:nth-child(1) > div > div.yuRUbf > a > div > cite";

                query.SendKeys("download  sample excel");

                query.Submit();

                Thread.Sleep(2000);

                driver.FindElement(By.CssSelector(excel_demo)).Click();

                


                //string userRoot = System.Environment.GetEnvironmentVariable("USERPROFILE");
                //string downloadFolder = Path.Combine(userRoot, "Documents");
                //Console.WriteLine(userRoot);
                //Open excel report
                //FileStream file2 = new FileStream(downloadFolder + "\\Export.xlsx", FileMode.Open, FileAccess.Read);
                //XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
                //ISheet sheet2 = workbook2.GetSheet("Sheet1");



                driver.Quit();
            }

        }
    }
}



