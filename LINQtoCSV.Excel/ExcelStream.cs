using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using LINQtoCSV;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LINQtoCSV.Excel
{
    internal class ExcelStream
    {

        private OpenXmlWriter _writer;
        private SpreadsheetDocument _spreadsheet;


        IExcelDataReader reader;
        String sheetName;

        // Current line number in the file. Only used when reading a file, not when writing a file.
        private int m_lineNbr = 0;
        public ExcelStream(Stream inStream, Stream outStream, String sheetName)
        {
            this.sheetName = sheetName;

            if (inStream != null)
            {
                reader = ExcelReaderFactory.CreateReader(inStream);

                do
                {
                    if (reader.Name != sheetName)
                    {
                        continue;
                    }
                    else
                    {
                        break;
                    }
                } while (reader.NextResult());

            }

            if (outStream != null)
            {
                _spreadsheet = SpreadsheetDocument.Create(outStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);


                //create workbook part
                WorkbookPart wbp = _spreadsheet.AddWorkbookPart();
                wbp.Workbook = new Workbook();
                Sheets sheets = wbp.Workbook.AppendChild<Sheets>(new Sheets());

                //create worksheet part, and add it to the sheets collection in workbook
                WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();
                Sheet sheet = new Sheet() { Id = _spreadsheet.WorkbookPart.GetIdOfPart(wsp), SheetId = 1, Name = sheetName };
                sheets.Append(sheet);

                _writer = OpenXmlWriter.Create(wsp);


            }

        }

        private Func<String, DocumentFormat.OpenXml.Spreadsheet.Cell> _getCell = new Func<string, DocumentFormat.OpenXml.Spreadsheet.Cell>(delegate (String value)
        {
            return new DocumentFormat.OpenXml.Spreadsheet.Cell()
            {
                DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value)
            };
        });

        public void StartWriteHead()
        {
            _writer.WriteStartElement(new Worksheet());
            _writer.WriteStartElement(new SheetData());
        }

        public void EndWriteHead()
        {
            _writer.WriteEndElement(); //end of SheetData
            _writer.WriteEndElement(); //end of worksheet
            _writer.Close();

            _spreadsheet.Dispose();

        }

        public void WriteRow(List<string> row)
        {
            _writer.WriteStartElement(new Row());

            foreach (var item in row)
            {
                _writer.WriteElement(_getCell(item));
            }
            _writer.WriteEndElement();
        }

        public bool ReadRow(IDataRow row, Dictionary<int, int> charactersLength = null)
        {
            row.Clear();

            m_lineNbr++;

            if (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    int? itemLength = null;
                    if(charactersLength != null && charactersLength.ContainsKey(i + 1))
                    {
                        itemLength = charactersLength[i + 1];
                    }

                    var value = Convert.ToString(reader.GetValue(i));

                    if (value != null && itemLength.HasValue && itemLength.Value  != 0 && value.Length > itemLength.Value)
                    {
                        value = value.Substring(0, itemLength.Value - 1);
                    }

                    row.Add(new DataRowItem(value, m_lineNbr));
                }

                return true;
            }
            else
            {
                return false;
            }


        }
    }
}
