using System;
using System.Data;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


namespace AlfaBank_TestCase_Proshunin
{
    class Program
    {
        private const string defaultString = "";
        readonly string rootProject = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName).FullName;
        private readonly string fileNameExcel = defaultString;
        private readonly string extension = defaultString;
        private readonly string url = defaultString;

        public Program(string fileNameExcel, string extension, string url)
        {
            this.fileNameExcel = fileNameExcel ?? throw new ArgumentNullException(nameof(fileNameExcel));
            this.extension = extension ?? throw new ArgumentNullException(nameof(extension));
            this.url = url ?? throw new ArgumentNullException(nameof(url));
        }

        public string FileNameExcel => fileNameExcel;

        public string Extension => extension;

        public string Url => url;

        static void Main(string[] args)
        {

            Program p = new Program(@"\challenge", ".xlsx", "http://www.rpachallenge.com/");

            string path = p.rootProject + p.FileNameExcel + p.Extension;

            DataTable extractedDT = DataSheet.GetDTFromExcel(path);

            IWebDriver chromeDriver = new ChromeDriver(Environment.CurrentDirectory);

            chromeDriver.Navigate().GoToUrl(p.Url);

            chromeDriver.Manage().Window.Maximize();

            chromeDriver.FindElement(By.XPath("//a[text()= 'Input Forms']")).Click();

            chromeDriver.FindElement(By.XPath("//button[text()= 'Start']")).Click();

            // XPath variable
            string templateXPath = "//input[@ng-reflect-name=";
            string labelFirstName = "'labelFirstName']";
            string labelLastName = "'labelLastName']";
            string labelCompanyName = "'labelCompanyName']";
            string labelRole = "'labelRole']";
            string labelAdress = "'labelAddress']";
            string labelEmail = "'labelEmail']";
            string labelPhone = "'labelPhone']";

            foreach (DataRow row in extractedDT.Rows)
            {

                //first name index 0
                chromeDriver.FindElement(By.XPath(templateXPath + labelFirstName)).SendKeys(row[0].ToString());
                //last name index 1
                chromeDriver.FindElement(By.XPath(templateXPath + labelLastName)).SendKeys(row[1].ToString());
                //company name index 2
                chromeDriver.FindElement(By.XPath(templateXPath + labelCompanyName)).SendKeys(row[2].ToString()); 
                //Role In Comp index 3
                chromeDriver.FindElement(By.XPath(templateXPath + labelRole)).SendKeys(row[3].ToString());
                //address index 4
                chromeDriver.FindElement(By.XPath(templateXPath + labelAdress)).SendKeys(row[4].ToString());
                //email index 5
                chromeDriver.FindElement(By.XPath(templateXPath + labelEmail)).SendKeys(row[5].ToString());
                //phone index 6
                chromeDriver.FindElement(By.XPath(templateXPath + labelPhone)).SendKeys(row[6].ToString());

                chromeDriver.FindElement(By.XPath("//input[@value='Submit']")).Click();
            }
        }
    }
}
