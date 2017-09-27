using LINQtoCSV.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

namespace LINQtoCSV.Excel.Tests
{
    [TestClass()]
    public class CsvContextWriteTests : Test
    {
        [TestMethod()]
        public void GoodFileCommaDelimitedNamesInFirstLineNLnl()
        {
            // Arrange

            List<ProductData> dataRows_Test = new List<ProductData>();
            dataRows_Test.Add(new ProductData { retailPrice = 4.59M, name = "Wooden toy", startDate = new DateTime(2008, 2, 1), nbrAvailable = 67 });
            dataRows_Test.Add(new ProductData { onsale = true, weight = 4.03, shopsAvailable = "Ashfield", description = "" });
            dataRows_Test.Add(new ProductData { name = "Metal box", launchTime = new DateTime(2009, 11, 5, 4, 50, 0), description = "Great\nproduct" });

            ExcelFileDescription fileDescription_namesNl2 = new ExcelFileDescription
            {
                FirstLineHasColumnNames = true,
                EnforceCsvColumnAttribute = false,
                TextEncoding = Encoding.Unicode,
                FileCultureName = "nl-Nl" // default is the current culture
            };
            
            string filePath = @"TestData\GoodFileCommaDelimitedNamesInFirstLineNLnl.xlsx";

            string sheetName = "Sheet1";

            // Act and Assert

            AssertWrite(dataRows_Test, filePath, sheetName , fileDescription_namesNl2);
        }
    }
}
