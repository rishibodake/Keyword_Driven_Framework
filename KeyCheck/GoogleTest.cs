using KeywordDriven;
using NUnit.Framework;

namespace KeyCheck
{
    public class Tests
    {
        string filePath = "";
        [SetUp]
        public void Setup()
        {
            filePath = @"C:\Users\abhib\source\repos\KeywordDriven\KeywordDriven\Excel File\first_keysheet.xlsx";
        }

        [Test]
        public void SearchTest()
        {
            ExcelAccess check = new ExcelAccess(filePath, 2);
            check.ExecutionEngine();

           // Assert.AreEqual("New", x.ToString());
        }
    }
}