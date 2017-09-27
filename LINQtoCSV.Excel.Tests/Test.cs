using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LINQtoCSV.Excel.Tests
{
    public abstract class Test
    {
        /// <summary>
        /// Takes a string and converts it into a StreamReader.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        protected Stream StreamReaderFromString(string s)
        {
            byte[] stringAsByteArray = System.Text.Encoding.UTF8.GetBytes(s);
            return new MemoryStream(stringAsByteArray);
            
        }

        protected void AssertCollectionsEqual<T>(IEnumerable<T> actual, IEnumerable<T> expected) where T : IAssertable<T>
        {
            int count = actual.Count();
            Assert.AreEqual(count, expected.Count(), "counts");

            for(int i = 0; i < count; i++)
            {
                actual.ElementAt(i).AssertEqual(expected.ElementAt(i));
            }
        }

        /// <summary>
        /// Used to test the Read method. 
        /// </summary>
        /// <typeparam name="T">
        /// Type of the output elements.
        /// </typeparam>
        /// <param name="testInput">
        /// String representing the contents of the file or StreamReader. This string is fed to the Read method
        /// as though it came from a file or StreamReader.
        /// </param>
        /// <param name="fileDescription">
        /// Passed to Read.
        /// </param>
        /// <returns>
        /// Output of Read.
        /// </returns>
        public IEnumerable<T> TestRead<T>(string filePath, string sheetName , ExcelFileDescription fileDescription) where T : class, new()
        {
            ExcelContext cc = new ExcelContext();
            return cc.Read<T>(filePath, sheetName, fileDescription);
        }

        /// <summary>
        /// Executes a Read and tests whether it outputs the expected records.
        /// </summary>
        /// <typeparam name="T">
        /// Type of the output elements.
        /// </typeparam>
        /// <param name="testInput">
        /// String representing the contents of the file or StreamReader. This string is fed to the Read method
        /// as though it came from a file or StreamReader.
        /// </param>
        /// <param name="fileDescription">
        /// Passed to Read.
        /// </param>
        /// <param name="expected">
        /// Expected output.
        /// </param>
        public void AssertRead<T>(string filePath, string sheetName, ExcelFileDescription fileDescription, IEnumerable<T> expected)
            where T : class, IAssertable<T>, new()
        {
            IEnumerable<T> actual = TestRead<T>(filePath, sheetName, fileDescription);
            AssertCollectionsEqual<T>(actual, expected);
        }

        /// <summary>
        /// Used to test the Write method
        /// </summary>
        /// <typeparam name="T">
        /// The type of the input elements.
        /// </typeparam>
        /// <param name="values">
        /// The collection of input elements.
        /// </param>
        /// <param name="fileDescription">
        /// Passed directly to write.
        /// </param>
        /// <returns>
        /// Returns a string with the content that the Write method writes to a file or TextWriter.
        /// </returns>
        public void TestWrite<T>(IEnumerable<T> values, string filePath, string sheetName, ExcelFileDescription fileDescription) where T : class
        {
            ExcelContext cc = new ExcelContext();
            cc.Write(values, filePath, sheetName, fileDescription);
        }

        /// <summary>
        /// Executes a Write and tests whether it outputs the expected records.
        /// </summary>
        /// <typeparam name="T">
        /// The type of the input elements.
        /// </typeparam>
        /// <param name="values">
        /// The collection of input elements.
        /// </param>
        /// <param name="fileDescription">
        /// Passed directly to write.
        /// </param>
        /// <param name="expected">
        /// Expected output.
        /// </param>
        public void AssertWrite<T>(IEnumerable<T> values, string filePath, string sheetName, ExcelFileDescription fileDescription) where T : class
        {
            TestWrite<T>(values, filePath, sheetName, fileDescription);
        }
    }
}
