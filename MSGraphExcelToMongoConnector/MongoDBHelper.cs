using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Driver;
using MongoDB.Driver.GridFS;
using MongoDB.Bson;

namespace MSExcelToMongoConnector
{
    public static class MongoDBHelper
    {
        private static IMongoClient _client;
        private static IMongoDatabase _database;

        // initialize the helper
        public static bool initialize(string connectionString, string database)
        {
            try {
                _client = new MongoClient(connectionString);
                _database = _client.GetDatabase(database);
                return true;
            } catch (Exception e)
            {
                Logger.logMessage("Cannot establish connection to MongoDB. Exception: " + e.Message, Logger.LogLevel.ERROR);
                return false;
            }
        }


        // refresh collection - drop and refresh collection with new set of data
        public static async Task<bool> refreshCollection(string collection, List<BsonDocument> data, int batchSize)
        {
            try { 
                // drop the collection
                await _database.DropCollectionAsync(collection);

                // populate the collection with data
                var col = _database.GetCollection<BsonDocument>(collection);
                var itemsToAdd = new List<WriteModel<BsonDocument>>();
                if (data.Count > batchSize)
                {
                    int itemsInBatch = 0;
                    int totalItemCount = 0;
                    foreach (var item in data)
                    {

                        var itemToAdd = new InsertOneModel<BsonDocument>(item);
                        itemsToAdd.Add(itemToAdd);
                        itemsInBatch++;
                        totalItemCount++;
                        if (itemsInBatch >= batchSize || totalItemCount == data.Count)
                        {
                            await col.BulkWriteAsync(itemsToAdd);
                            itemsInBatch = 0; // reset batch
                            itemsToAdd.Clear();
                        }
                    }
                } else { // data size is smaller than batch
                    foreach (var item in data)
                    {
                        var itemToAdd = new InsertOneModel<BsonDocument>(item);
                        itemsToAdd.Add(itemToAdd);
                    }
                    await col.BulkWriteAsync(itemsToAdd);
                }
            } catch (Exception e)
            {
                Logger.logMessage("Cannot refresh collection in MongoDB. Exception: " + e.Message, Logger.LogLevel.ERROR);
                return false;
            }
            return true;
            
        }



        public static async Task<bool> uploadFileToGridFS(string filename)
        {
            var bucket = new GridFSBucket(_database, new GridFSBucketOptions
            {
                BucketName = "excelSpreadsheets",
                ChunkSizeBytes = 1048576, // 1MB
                WriteConcern = WriteConcern.W1,
                ReadPreference = ReadPreference.Primary
            }
            );
            try
            {
                byte[] source = System.IO.File.ReadAllBytes(filename);
                await bucket.UploadFromBytesAsync(filename, source);
            }
            catch (Exception e)
            {
                Logger.logMessage("Cannot establish connection to MongoDB. Exception: " + e.Message, Logger.LogLevel.ERROR);
                return false;
            }


            return true;
        }

        //TODO: download raw file

        //mainly for tests
        public static async Task<bool> verifyCollectionItemCount(string collection, long itemCount)
        {
            try
            {
                // populate the collection with data
                var col = _database.GetCollection<BsonDocument>(collection);
                var actualItemCount = await col.CountAsync(new BsonDocument());
                if (actualItemCount != itemCount)
                {
                    return false;
                } else
                {
                    return true;
                }
            }
            catch (Exception e)
            {
                //throw (e);
                return false;
            }
        }

    }

}
