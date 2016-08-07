using System;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MSExcelToMongoConnector;
using MongoDB.Bson;
using System.Collections.Generic;
using System.Data;

namespace MSExcelToMongoConnectorTest
{
    [TestClass]
    public class MongoHelperTests
    {
        [TestMethod]
        public void TestConnection()
        {
            var connectionString = "mongodb://localhost:27017";
            var database = "MSGraphConnector";
            var success = MongoDBHelper.initialize(connectionString, database);
            Assert.IsTrue(success);
        }

        [TestMethod]
        public void TestConnectionShouldFail()
        {
            var connectionString = "mongodb1://localhost:33333?connect=direct";
            var database = "MSGraphConnector";
            var success = MongoDBHelper.initialize(connectionString, database);
            Assert.IsFalse(success, "couldn't connect to DB");
        }

        [TestMethod]
        public async Task TestLoadDocs()
        {
            var connectionString = "mongodb://localhost:27017";
            var database = "MSGraphConnector";
            MongoDBHelper.initialize(connectionString, database);
            var collection = "DummySpreadsheet";
            var testDocs = new List<BsonDocument>();
            testDocs.Add(new BsonDocument("col1", "value1").Add("col2", "value2"));
            testDocs.Add(new BsonDocument("col1", "value2").Add("col2", "value3"));
            testDocs.Add(new BsonDocument("col1", "value2").Add("col2", "value1"));
            testDocs.Add(new BsonDocument("col1", "value2").Add("col2", "value2"));

            var success = await MongoDBHelper.refreshCollection(collection, testDocs, 5);
            Assert.IsTrue(success, "couldn't refresh dummy collection");
            success = await MongoDBHelper.verifyCollectionItemCount(collection, 4);
            Assert.IsTrue(success, "got wrong number of items");
        }
    }

    [TestClass]
    public class ExcelHelperTests
    {
        [TestMethod]
        public void GetSpreadsheetData()
        {
            IEnumerable<DataTable> excelDataTables = MSExcelHelper.getWorksheetData(@"c:\temp\test.xlsx");
            int i = 0;
            foreach (var table in excelDataTables)
            {
                i++;
            }
            Assert.AreEqual(i, 1);

        }
    }
}
