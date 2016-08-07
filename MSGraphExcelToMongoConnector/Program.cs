using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using MongoDB.Bson;

namespace MSExcelToMongoConnector
{
    class Program
    {
        static void Main(string[] args)
        {
            Logger.initialize();
            Logger.logMessage("Program Started", Logger.LogLevel.INFO);
            if (args.Length == 0)
            {
                showSyntax();
            }
            else
            {
                YamlConfiguration.readConfig();
                MainAsync(args).Wait();
            }
        }

        private static void showSyntax()
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Usage: MSExcelToMongoConnector.exe [file path]");
            Console.ResetColor();
        }

        public static async Task MainAsync(string[] args)
        {

            var connectionString = ConnectorConfiguration.connectionString; 
            var database = ConnectorConfiguration.databaseName;
            var success = MongoDBHelper.initialize(connectionString, database);
            if (success == false)
            {
                Logger.logMessage("Cannot establish connection to MongoDB. Quitting" , Logger.LogLevel.ERROR);
                System.Threading.Thread.CurrentThread.Abort();
            }
            var fileName = args[0];
            Logger.logMessage("Reading from spreadsheet", Logger.LogLevel.INFO);
            IEnumerable<DataTable> excelDataTables = MSExcelHelper.getWorksheetData(fileName); 
            if (excelDataTables == null)
            {
                Logger.logMessage("Couldn't get data from spreadsheet", Logger.LogLevel.ERROR);
                System.Threading.Thread.CurrentThread.Abort();
            }
            Logger.logMessage("Got data from excel", Logger.LogLevel.INFO);
            
            if (ConnectorConfiguration.storeRaw == true) { 
                Logger.logMessage("Getting data to MongoDB using GridFS", Logger.LogLevel.INFO);
                await MongoDBHelper.uploadFileToGridFS(fileName);
            }

            foreach (var table in excelDataTables) // a table is a worksheet
            {
                Logger.logMessage("Reading spreadsheet: " + table.TableName, Logger.LogLevel.INFO);
                List<BsonDocument> dataToMongoDB = new List<BsonDocument>();
                int i = 1;
                foreach (DataRow row in table.Rows)
                {
                    var doc = new BsonDocument();
                    foreach (DataColumn col in table.Columns)
                    {
                        
                        doc.Add(col.ColumnName.ToString(), row[col.ColumnName.ToString()].ToString());
                    }
                    Logger.logMessage("Added document #" + i.ToString(), Logger.LogLevel.INFO);
                    dataToMongoDB.Add(doc);
                    i++;
                }

                // now send through the list to MongoDB; 
                Logger.logMessage("Sending data to MongoDB", Logger.LogLevel.INFO);
                
                await MongoDBHelper.refreshCollection(removeIllegalChars(System.IO.Path.GetFileName(fileName) + "." + table.TableName), dataToMongoDB, ConnectorConfiguration.batchSize);
                Logger.logMessage("Finished reading spreadsheet " + table.TableName, Logger.LogLevel.INFO);
            }
        }

        private static string removeIllegalChars(string raw)
        {
            return raw.Replace(" ", "").Replace("$", "");
        }
        
    }
}
