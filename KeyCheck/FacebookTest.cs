using KeywordDriven;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace KeyCheck
{
    public class FacebookTest
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
