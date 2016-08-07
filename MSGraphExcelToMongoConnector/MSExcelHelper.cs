using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel.Core;
using Excel;
using System.IO;
using System.Data;

namespace MSExcelToMongoConnector
{
    public class MSExcelHelper
    {
        public static IEnumerable<DataTable> getWorksheetData(string filename)
        {
            try {
                // read file into a dataset
                FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet result = excelReader.AsDataSet();

                var tables = from table in result.Tables.Cast<DataTable>() select table;
                return tables;
                
                
            } catch (Exception e) {
                Logger.logMessage("Couln't get spreadsheet data. Exception: " + e.Message, Logger.LogLevel.ERROR);
                return null;
            }
        }
    }
}
