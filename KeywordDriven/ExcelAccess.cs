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
            
            int k = 1;
            for(int index = 0; index < 11; index++)
            {

                Range range = (Range)workSheet.Cells[index+1,k+1];
                string locator = range.Value.ToString();


                if (!locator.Equals("NA") && locator.Contains("="))
                {
                    locatorName = locator.Split('=')[0] ;//name
                    locatorValue = locator.Split('=')[1];//q
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
