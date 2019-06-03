using System;
using System.Collections;
using System.Text;
using Microsoft.SqlServer.Server;
using System.Data.SqlTypes;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Data.SQLite;

namespace Synchronizers{

    public class TableSynchronizer
    {

        string sourceServer = "";
        string sourceUser = "";
        string sourcePassword = "";
        string sourceDatabase = "";
        string sourceTable = "";
        SqlConnection sourceConnection;

        string destinationServer = "";
        string destinationUser = "";
        string destinationPassword = "";
        string destinationDatabase = "";
        string destinationTable = "";
        SqlConnection destinationConnection;
        string tabDiffCmd = @"C:\Program Files\Microsoft SQL Server\100\COM\tablediff.exe";
        string commandString = "";
        string sqlCompSqlFile;

        public TableSynchronizer()
        {

        }

        public string getSourceServer()
        {
            return this.sourceServer;
        }

        public string getsourceUser()
        {
            return this.sourceUser;
        }
        public string getsourcePassword()
        {
            return this.sourcePassword;
        }
        public string getsourceDatabase()
        {
            return this.sourceDatabase;
        }
        public string getsourceTable()
        {
            return this.sourceTable;
        }

        public string getDestinationServer()
        {
            return this.destinationServer;
        }

        public string getDestinationUser()
        {
            return this.destinationUser;
        }
        public string getDestinationPassword()
        {
            return this.destinationPassword;
        }
        public string getDestinationDatabase()
        {
            return this.destinationDatabase;
        }

        public string getSQLFile()
        {
            return this.sqlCompSqlFile;
        }

        public string getDestinationTable()
        {
            return this.destinationTable;
        }
        public void setSourceServer(string server)
        {
            this.sourceServer = server;
        }

        public void setsourceUser(string user)
        {
            this.sourceUser = user;
        }
        public void setsourcePassword(string password)
        {
            this.sourcePassword = password;
        }
        public void setsourceDatabase(string database)
        {
            this.sourceDatabase = database;
        }
        public void setsourceTable(string table)
        {
            this.sourceTable = table;
        }

        public void setDestinationServer(string server)
        {
            this.destinationServer = server;
        }

        public void setDestinationUser(string user)
        {
            this.destinationUser = user;
        }
        public void setDestinationPassword(string password)
        {
            this.destinationPassword = password;
        }
        public void setDestinationDatabase(string database)
        {
            this.destinationDatabase = database;
        }
        public void setDestinationTable(string table)
        {
            this.destinationTable = table;
        }
        public void setConnectionMode(int mode)
        {
            this.connectionMode = mode;

        }


        public void getSQLFile(string fileName)
        {
            this.sqlCompSqlFile = fileName;
        }

        public string getCommandString()
        {
            tabDiffCmd = File.Exists(tabDiffCmd) ? tabDiffCmd : @"C:\Program Files\Microsoft SQL Server\110\COM\tablediff.exe";
            if (File.Exists(tabDiffCmd))
            {
                commandString = @"" + tabDiffCmd + "  -sourceserver " + this.getSourceServer() + " -sourceuser " + this.getSourceUser() + " -sourcepassword " + this.getSourcePassword() + " -sourcedatabase " + this.getSourceData() + " -sourcetable " + this.getSourceTable() +
                                "-destinationserver " + this.getDestinationServer() + " -destinationuser " + this.getDestinationUser() + " -destinationpassword " + this.getDestinationPassword() + "-destinationdatabase " + this.getDestinationDatabase() + " -destinationtable " + this.getDestinationTable() + "  -f " + sqlCompSqlFile;
            }
            else
            {

                Console.WriteLine("tablediff command not found");
            }

            return tabDiffCmd;
        }

        public void runTableComparison(string command)
        {

            try
            {

                Process cmd = new Process();
                cmd.StartInfo.FileName = this.getCommandString();
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.StartInfo.UseShellExecute = false;
                cmd.Start();
                cmd.StandardInput.Close();
                cmd.WaitForExit();
                string result = proc.StandardOutput.ReadToEnd();
                Console.WriteLine(result);

            }
            catch (Exception e)
            {
                Console.WriteLine("Error running table comparison: " + e.message);
                e.printStackTrace();
            }
        }

        public void initConnections()
        {
            try
            {
                int sessionConnectionMode = this.getConnectionMode();
                sourceConnection = new SqlConnection("Network Library=DBMSSOCN;Data Source=" + this.getSourceServer() + ",1433;database=" + this.getSourceDatabase() + ";User id=" + this.getSourceUser() + ";Password=" + this.getsourcePassword() + ";Connection Timeout=0");
                sourceConnection.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error running table comparison: " + e.message);
                e.printStackTrace();

            }

        }
        public string runSyncSQL(string queryFile)
        {
            string sql_query = "";
            queryFile = queryFile != null ? getSQLFile() : "";
            try
            {
                if (File.Exists(queryFile))
                {

                    sql_query = File.ReadAllText(queryFile);
                    SqlCommand cmd = new SqlCommand(sql_query, sourceConnection);
                    cmd.CommandTimeout = 0;
                    Console.WriteLine("Running SQL query: " + sql_query);
                    SqlDataReader dr = cmd.ExecuteNonQuery();

                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error running query: " + e.message);
                e.printStackTrace();

            }
        }
    }
        public class DatabaseSynchronizer {


        public DatabaseSynchronizer() {


        }

            public void RunDataSynchronizer
            {






            }



        }

    }
  

  
