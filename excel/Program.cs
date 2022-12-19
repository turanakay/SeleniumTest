using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ExcelApp = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Data.OleDb;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace excel
{
    class Program
    {


        static void Main(string[] args)
        {

            //Create COM Objects.
            ExcelApp.Application excelApp = new ExcelApp.Application();
            DataRow myNewRow;
            DataTable myTable;


            if (excelApp == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }

            //
            ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\Turan\Desktop\user2.xlsx");
            ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
            ExcelApp.Range excelRange = excelSheet.UsedRange;

            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            //Set DataTable Name and Columns Name
            myTable = new DataTable("MyDataTable");
            myTable.Columns.Add("name", typeof(string));
            myTable.Columns.Add("password", typeof(string));
            myTable.Columns.Add("event", typeof(string));



            //first row using for heading, start second row for data
            for (int i = 2; i <= rows; i++)
            {
                myNewRow = myTable.NewRow();
                myNewRow["name"] = excelRange.Cells[i, 1].Value2.ToString(); //string
                myNewRow["password"] = excelRange.Cells[i, 2].Value2.ToString(); //string
                myNewRow["event"] = excelRange.Cells[i, 3].Value2.ToString(); //string

                myTable.Rows.Add(myNewRow);
            }
            int try_count = 1;
            foreach (DataRow j in myTable.Rows)
            {

                string name = j[0].ToString();
                string pass = j[1].ToString();
                string event_name = j[2].ToString();
                Console.WriteLine("Deneme " + try_count + " Başlıyor. ");
                var options = new ChromeOptions();
                var driver = new ChromeDriver("C:\\Users\\Turan\\Desktop\\d\\", options);
                


                driver.Url = "https://www.passo.com.tr/tr";

                Task.Delay(1000).Wait();
                driver.FindElement(By.XPath("/html/body/div/div/a")).Click();
                Thread.Sleep(1000);
                driver.Manage().Window.Maximize();
                Thread.Sleep(3000);
                driver.ExecuteScript("window.scrollBy(0,550)");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-home/app-search/section[2]/event-search-and-list/div[1]/div[1]/div[1]/div[1]/div")).Click();
                Thread.Sleep(3000);
                driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-home/app-search/section[2]/event-search-and-list/div[1]/div[1]/div[1]/div[1]/div/div/div[2]/div[1]/span[1]")).Click();
                Thread.Sleep(100000);
                //driver.Quit();

            
                Task.Delay(1000).Wait();
               



               


                bool pop_up = false;
                try
                {

                    pop_up = driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/button[2]")).Displayed;
                    pop_up = true;
                }
                catch (Exception)
                {

                    pop_up = false;
                }
                if (pop_up == true)
                {
                    driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/button[2]")).Click();
                }
                else
                {
                    Console.WriteLine("pop_upı kim kapattı?");
                }



                Console.WriteLine("BAŞARILI BİR GİRİŞİM  OLDU.");
   
                Task.Delay(1000).Wait();
             
                driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-header/div[1]/div/div[1]/input")).SendKeys(event_name);
              
                Task.Delay(3000).Wait();
            
                
                Task.Delay(2000).Wait();
                bool element = false;
                try
                {
                    
                    element = driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-header/div[1]/nav/div/div[2]/ul/li")).Displayed;
                    element = true;
                }
                catch (Exception)
                {

                    element = false;
                }

                if (element == false)
                {
                    Console.WriteLine("Yanlış etkinlik adı girdiniz. Terminating...");
                    Task.Delay(1000).Wait();
                    driver.Close();
                    driver.Quit();

                }
                else
                {
                    Console.WriteLine("ETKİNLİK BULUNDU");
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-header/div[1]/div/div[1]/ul/li")).Click();

                    Task.Delay(1000).Wait();
                    bool pop_up2 = false;
                    try
                    {
                        pop_up2 = driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/button[2]")).Displayed;
                        pop_up2 = true;
                    }
                    catch (Exception)
                    {

                        pop_up2 = false;
                    }
                    if (pop_up2 == true)
                    {
                        driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/button[2]")).Click();
                    }
                    else
                    {
                        Console.WriteLine("pop_upı kim kapattı?");
                    }


                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-event/section[1]/div/div/div/div[1]/div[2]/div[3]/button")).Click();
                    Task.Delay(1000).Wait();

                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-login/section/div/div/div/div/div[2]/div/div/div[1]/div/quick-form/div/quick-input[1]/input")).SendKeys(name);
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-login/section/div/div/div/div/div[2]/div/div/div[1]/div/quick-form/div/quick-input[2]/input")).SendKeys(pass);
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-login/section/div/div/div/div/div[2]/div/div/div[1]/div/quick-form/div/div[2]/button[2]")).Click();
                    Task.Delay(1000).Wait();

                    Task.Delay(1000).Wait();

                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-seat/div/div[3]/div/div[2]/div[3]/div[3]/div/div[2]/div/div/div/select/option[2]")).Click();

                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-seat/div/div[3]/div/div[2]/div[4]/select/option[2]")).Click();
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-seat/div/div[3]/div/div[2]/div[5]/div[3]/div[1]/button")).Click();
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-basket/section/div/div[6]/button[4]")).Click();

                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-event-group-payment/section/div/div[6]/app-delivery/div/div/quick-form/div/quick-select/ng-select/div/span")).Click();
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-event-group-payment/section/div/div[6]/app-delivery/div/div/quick-form/div/quick-select/ng-select/ng-dropdown-panel/div[2]/div[2]/div/span")).Click();
                    Console.WriteLine("PDF bilet seçildi");
                    Task.Delay(1000).Wait();

                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-event-group-payment/section/div/div[6]/app-delivery/div/div[1]/quick-form/div/div[3]/button[2]")).Click();
                    Task.Delay(1000).Wait();

                    driver.Manage().Window.Maximize();
                    driver.ExecuteScript("window.scrollBy(0,-250)");
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-event-group-payment/section/div/div[6]/div[3]/quick-form/div/div[1]/quick-checkbox/div/div/label/span")).Click();
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-event-group-payment/section/div/div[6]/div[3]/quick-form/div/div[2]/quick-checkbox/div/div/label/span")).Click();

                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-event-group-payment/section/div/div[6]/div[3]/quick-form/div/quick-checkbox/div/div/label/span")).Click();
                    Console.WriteLine("Sözleşmeler imzalandı.");
                    //Thread.Sleep(1000);
                    driver.FindElement(By.XPath("/html/body/app-root/app-layout/app-event-group-payment/section/div/div[6]/div[3]/quick-form/div/div[4]/button[2]")).Click();
                    Console.WriteLine("Bilet alındı.");
                    Task.Delay(1000).Wait();
                    driver.Quit();


                }
                try_count++;







            }
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.ReadLine();
            Console.WriteLine("Console kapanıyor...");
            Thread.Sleep(10000);
            Environment.Exit(0);


        }
 
    }

}
