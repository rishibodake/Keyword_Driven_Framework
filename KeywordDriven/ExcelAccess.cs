using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using _Excel = Microsoft.Office.Interop.Excel;

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
            for(int index = 0; index < 4; index++)
            {

                Range range = (Range)workSheet.Cells[index+1,k+1];
                string locator = range.Value.ToString();


                if (!locator.Equals("NA"))
                {
                    locatorName = locator.Split('=')[0] ;//name
                    //locatorValue = cellValue.Split('=')[1];//q
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
                    case "quit":
                        driver.Quit();
                        break;
                    default:
                        break;
                }

                switch (locatorName)
                {
                    case "name":
                        IWebElement ele = driver.FindElement(By.Name("q"));
                        if (action_values.Equals("sendkeys"))
                        {
                            ele.SendKeys(values + Keys.Enter);
                        }
                        else if(action_values.Equals("click"))
                        {
                            ele.Click();
                        }
                        break;
                    default:
                        break;
                }                              
            }
        }
    }
}
