        using System;
        using System.Data;
        using System.Data.SQLite;
        using ExcelDna.Integration;

        namespace UsingSQLite
        {
            public static class MyFunctions
            {
                static SQLiteConnection _connection;
                static SQLiteCommand _productNameCommand;

                private static void EnsureConnection()
                {
                    if (_connection == null)
                    {
                        _connection = new SQLiteConnection(@"Data Source=C:\Temp\Northwind.db");
                        _connection.Open();

                        _productNameCommand = new SQLiteCommand("SELECT ProductName FROM Products WHERE ProductID = @ProductID", _connection);
                        _productNameCommand.Parameters.Add("@ProductID", DbType.Int32);
                    }
                }

                public static object ProductName(int productID)
                {
                    try
                    {
                        EnsureConnection();
                        _productNameCommand.Parameters["@ProductID"].Value = productID;
                        return _productNameCommand.ExecuteScalar();
                    }
                    catch (Exception ex)
                    {
                        return ex.ToString();
                    }
                }

            }
        }
