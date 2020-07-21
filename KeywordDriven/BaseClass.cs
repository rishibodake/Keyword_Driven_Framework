using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;

namespace KeywordDriven
{
    public class BaseClass
    {
        IWebDriver driver;
        public IWebDriver InitDriver(string browser)
        {
            switch (browser)
            {
                case "chrome":
                    ChromeOptions chromeOptions = new ChromeOptions();
                    chromeOptions.AddArgument("--disable-notifications");
                    driver = new ChromeDriver(chromeOptions);
                    driver.Manage().Window.Maximize();
                    break;
                case "firefox":
                    FirefoxOptions firefoxOptions = new FirefoxOptions();
                    firefoxOptions.SetPreference("dom.webnotifications.enabled", false);
                    driver = new FirefoxDriver(firefoxOptions);
                    break;
                case "edge":
                    driver = new EdgeDriver();
                    break;
                case "ie":
                    driver = new InternetExplorerDriver();
                    break;
                default:
                    driver = new ChromeDriver();
                    break;
            }
            return driver;
        }
    }
}
