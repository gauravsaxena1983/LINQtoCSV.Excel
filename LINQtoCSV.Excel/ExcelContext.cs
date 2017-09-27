using ExcelDataReader;
using LINQtoCSV;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LINQtoCSV.Excel
{
    public class ExcelContext
    {
        public IEnumerable<T> Read<T>(string fileName, String sheetName, ExcelFileDescription fileDescription) where T : class, new()
        {
            // Note that ReadData will not be called right away, but when the returned 
            // IEnumerable<T> actually gets accessed.

            return ReadData<T>(fileName, null, sheetName, fileDescription);
        }

        public IEnumerable<T> Read<T>(Stream stream, String sheetName) where T : class, new()
        {
            return Read<T>(stream, sheetName,  new ExcelFileDescription());
        }

        public IEnumerable<T> Read<T>(string fileName, String sheetName) where T : class, new()
        {
            return Read<T>(fileName, sheetName,  new ExcelFileDescription());
        }

        public IEnumerable<T> Read<T>(Stream stream,  string sheetName, ExcelFileDescription fileDescription) where T : class, new()
        {
            return ReadData<T>(null, stream, sheetName, fileDescription);
        }


        private IEnumerable<T> ReadData<T>(
                    string fileName,
                    Stream stream,
                    String sheetName,
                    ExcelFileDescription fileDescription) where T : class, new()
        {
            // If T implements IDataRow, then we're reading raw data rows 
            bool readingRawDataRows = typeof(IDataRow).IsAssignableFrom(typeof(T));

            // The constructor for FieldMapper_Reading will throw an exception if there is something
            // wrong with type T. So invoke that constructor before you open the file, because if there
            // is an exception, the file will not be closed.
            //
            // If T implements IDataRow, there is no need for a FieldMapper, because in that case we're returning
            // raw data rows.
            FieldMapper_Reading<T> fm = null;

            if (!readingRawDataRows)
            {
                fm = new FieldMapper_Reading<T>(fileDescription, fileName, false);
            }


            // -------
            // Each time the IEnumerable<T> that is returned from this method is 
            // accessed in a foreach, ReadData is called again (not the original Read overload!)
            //
            // So, open the file here, or rewind the stream.

            bool readingFile = !string.IsNullOrEmpty(fileName);

            if (readingFile)
            {
                stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
            }
            else
            {
                // Rewind the stream

                if ((stream == null) || (!stream.CanSeek))
                {
                    throw new BadStreamException();
                }
            }
            
            ExcelStream es = new ExcelStream(stream, null, sheetName);

            // If we're reading raw data rows, instantiate a T so we return objects
            // of the type specified by the caller.
            // Otherwise, instantiate a DataRow, which also implements IDataRow.
            IDataRow row = null;
            if (readingRawDataRows)
            {
                row = new T() as IDataRow;
            }
            else
            {
                row = new DataRow();
            }

            AggregatedException ae =
                new AggregatedException(typeof(T).ToString(), fileName, fileDescription.MaximumNbrExceptions);

            try
            {
                Dictionary<int, int> charLengths = null;
                if (!readingRawDataRows)
                {
                    charLengths = fm.GetCharLengths();
                }

                bool firstRow = true;
                while (es.ReadRow(row, charLengths))
                {
                    // Skip empty lines.
                    // Important. If there is a newline at the end of the last data line, the code
                    // thinks there is an empty line after that last data line.
                    if ((row.Count == 1) &&
                        ((row[0].Value == null) ||
                         (string.IsNullOrEmpty(row[0].Value.Trim()))))
                    {
                        continue;
                    }

                    if (firstRow && fileDescription.FirstLineHasColumnNames)
                    {
                        if (!readingRawDataRows) { fm.ReadNames(row); }
                    }
                    else
                    {
                        T obj = default(T);
                        try
                        {
                            if (readingRawDataRows)
                            {
                                obj = row as T;
                            }
                            else
                            {
                                obj = fm.ReadObject(row, ae);
                            }
                        }
                        catch (AggregatedException ae2)
                        {
                            // Seeing that the AggregatedException was thrown, maximum number of exceptions
                            // must have been reached, so rethrow.
                            // Catch here, so you don't add an AggregatedException to an AggregatedException
                            throw ae2;
                        }
                        catch (Exception e)
                        {
                            // Store the exception in the AggregatedException ae.
                            // That way, if a file has many errors leading to exceptions,
                            // you get them all in one go, packaged in a single aggregated exception.
                            ae.AddException(e);
                        }

                        yield return obj;
                    }
                    firstRow = false;
                }
            }
            finally
            {
                if (readingFile)
                {
                    stream.Close();
                }

                // If any exceptions were raised while reading the data from the file,
                // they will have been stored in the AggregatedException ae.
                // In that case, time to throw ae.
                ae.ThrowIfExceptionsStored();

                
            }

        }

        public void Write<T>(
            IEnumerable<T> values,
            string fileName,
            String sheetName,
            ExcelFileDescription fileDescription)
        {
            using (Stream sw = File.Open(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                WriteData<T>(values, fileName, sw, sheetName, fileDescription);
            }
        }

        public void Write<T>(
            IEnumerable<T> values,
            Stream stream, 
            String sheetName)
        {
            Write<T>(values, stream, sheetName, new ExcelFileDescription());
        }

        public void Write<T>(
            IEnumerable<T> values,
            string fileName,
            String sheetName)
        {
            Write<T>(values, fileName, sheetName, new ExcelFileDescription());
        }

        public void Write<T>(
            IEnumerable<T> values,
            Stream stream,
            String sheetName,
            ExcelFileDescription fileDescription)
        {
            WriteData<T>(values, null, stream, sheetName,  fileDescription);
        }

        private void WriteData<T>(
            IEnumerable<T> values,
            string fileName,
            Stream stream,
            String sheetName,
            ExcelFileDescription fileDescription)
        {
            FieldMapper<T> fm = new FieldMapper<T>(fileDescription, fileName, true);
            ExcelStream es = new ExcelStream(null, stream, sheetName);

            List<string> row = new List<string>();

            es.StartWriteHead();

            // If first line has to carry the field names, write the field names now.
            if (fileDescription.FirstLineHasColumnNames)
            {
                fm.WriteNames(row);
                es.WriteRow(row);
            }

            // -----

            foreach (T obj in values)
            {
                // Convert obj to row
                fm.WriteObject(obj, row);
                es.WriteRow(row);
            }

            es.EndWriteHead();

        }
    }
}
