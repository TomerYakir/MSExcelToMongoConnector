## Synopsis

MSExcelToMongoConnector - imports **Excel** (xlsx) spreadsheets to **MongoDB** 

## Prerequisites
To run this tool, you need .Net FW 4.5+

## Running the tool
Syntax:
MSExcelToMongoConnector.exe [file name]

Example:
MSExcelToMongoConnector.exe c:\temp\myFile.xlsx

## Configuration
Use the MSExcelToMongoConnector.yaml to configure the following:
**connectionString** - MongoDB connection string
**databaseName** - MongoDB database name to use
**batchSize** - batch size when storing the data in MongoDB
**storeRaw** - true/false - indicates whether to store the raw Excel files in the database (using GridFS)

Sample configuration:
connectionString: mongodb://localhost:27017,localhost:27018,localhost:27019?replicaSet=test
databaseName: MSExcelConnector
batchSize: 10
storeRaw: true


## Contributors
Tomer Yakir, MongoDB

## License and Packages
This tool uses the following packages:
ExcelDataReader
Log4Net
MongoDB .Net Driver
YamlDotNet
