using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace MSExcelToMongoConnector
{
    public static class YamlConfiguration
    {
        public static void readConfig()
        {
            var filename = "MSExcelToMongoConnector.yaml";
            string fileContent = File.ReadAllText(filename);
            var input = new StringReader(fileContent);
            var deserializer = new Deserializer();
            var config = deserializer.Deserialize<ConnectorInstanceConfiguration>(input);
            //TODO: work with another YAML deserializer that can work with static classes
            ConnectorConfiguration.connectionString = config.connectionString;
            ConnectorConfiguration.databaseName = config.databaseName;
            ConnectorConfiguration.batchSize = config.batchSize;
            ConnectorConfiguration.storeRaw = config.storeRaw;

        }
    }

    public static class ConnectorConfiguration
    {
        public static string connectionString { get; set; }
        public static string databaseName { get; set; }
        public static int batchSize { get; set; }
        public static bool storeRaw { get; set; }
    }

    public class ConnectorInstanceConfiguration
    {
        public string connectionString { get; set; }
        public string databaseName { get; set; }
        public int batchSize { get; set; }
        public bool storeRaw { get; set; }

    }

}
