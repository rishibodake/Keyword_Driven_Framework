using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using _Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Threading;

namespace KeywordDriven
{
    public class ExcelAccess
    {
        IWebDriver driver;
        

        _Application excel = new _Excel.Application();
        Workbook workBook;
        Worksheet workSheet;

        

        string locatorName = null;
        string locatorValue = null;

        public ExcelAccess(string path, int sheetNumber)
        {           
            workBook = excel.Workbooks.Open(path);
            workSheet = workBook.Worksheets[sheetNumber];
        }

        public void ExecutionEngine()
        {
            _Excel.Range last = workSheet.Cells.SpecialCells(_Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            _Excel.Range range12 = workSheet.get_Range("A1", last);

            int lastUsedRow = last.Row;
            int lastUsedColumn = last.Column;

            int k = 1;
            for(int index = 0; index < lastUsedRow; index++)
            {

                Range range = (Range)workSheet.Cells[index+1,k+1];
                string locator = range.Value.ToString();


                if (!locator.Equals("NA") && locator.Contains("="))
                {
                    locatorName = locator.Split(new[] { '=' }, 2)[0].Trim() ;//name
                    locatorValue = locator.Split(new[] { '=' }, 2)[1].Trim();//q
                }

                Range range1 = (Range)workSheet.Cells[index+1,k+2];
                string action_values = range1.Value.ToString();

                Range range2 = (Range)workSheet.Cells[index+1, k + 3];
                string values = range2.Value.ToString();


                switch (action_values)
                {
                    case "open browser":
                        BaseClass init = new BaseClass();
                        driver = init.InitDriver(values);                       
                        break;
                    case "enter url":
                        driver.Url = values;
                        //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                        break;
                    case "close":
                        driver.Close();
                        break;
                    case "quit":
                        driver.Quit();
                        break;
                    default:
                        break;
                }
              
                switch (locatorName)
                {
                    
                    case "name":
                        IWebElement elementByName = driver.FindElement(By.Name(locatorValue));                  
                        if (action_values.Equals("sendkeys"))
                        {
                            elementByName.SendKeys(values);
                        }
                        else if(action_values.Equals("click"))
                        {
                            elementByName.Click();
                        }
                        locatorName = null;
                        break;
                    case "id":
                        IWebElement elementById = driver.FindElement(By.Id(locatorValue));
                        if (action_values.Equals("sendkeys"))
                        {
                            elementById.SendKeys(values);                           
                        }
                        else if (action_values.Equals("click"))
                        {
                            elementById.Click();
                           // Thread.Sleep(3000);
                        }
                        locatorName = null;
                        break;
                    case "xpath":
                        IWebElement elementByXpath = driver.FindElement(By.XPath(locatorValue));
                        if (action_values.Equals("sendkeys"))
                        {
                            elementByXpath.SendKeys(values);
                            //Thread.Sleep(3000);
                        }
                        else if (action_values.Equals("click"))
                        {
                            elementByXpath.Click();
                            //Thread.Sleep(3000);
                        }
                        locatorName = null;
                        break;
                    case "css selector":
                        IWebElement elementByCssSelector = driver.FindElement(By.CssSelector(locatorValue));
                        if (action_values.Equals("sendkeys"))
                        {
                            elementByCssSelector.SendKeys(values);
                        }
                        else if (action_values.Equals("click"))
                        {
                            elementByCssSelector.Click();
                        }
                        locatorName = null;
                        break;
                    default:
                        break;
                }                              
            }
        }
    }
}
