using System;
using System.Collections;
using System.Text;
using Microsoft.SqlServer.Server;
using System.Data.SqlTypes;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace Synchronizers{

   
        public class DatabaseSynchronizer {
			
			
			string sourceServer = "";
			string sourceDatabase = "";
			SqlConnection destinationConnection;
			string destinationServer = "";
			string destinationDatabase = "";
		//	string commandString = "";
			ArrayList destinationTableList= new ArrayList();
			ArrayList syncedTableList= new ArrayList();
			ConnectionProperty conProps;
			Thread[] syncThreads;


			public DatabaseSynchronizer() {
				


			}

			 public int checkRunningThreadCount(){
				int runCount =-1;
				 for(int i=0; i< syncThreads.Length; i++){
					 if(null==syncThreads[i] || !syncThreads[i].IsAlive){
						 return i;
					 } else{
						 ++runCount; 
					 }
					 
				 }
							  return runCount;
				
			}		

   public void synchTables(string tableName){
	    Console.WriteLine("Synchronizing table: "+tableName);
	    new TableSynchronizer(conProps.getSourceServer(),conProps.getSourceDatabase(),tableName,conProps.getDestinationServer(),conProps.getDestinationDatabase(),tableName);
	      
   }
   
      public void synchTables(string tableName, string tabDiffCmd){
		 lock(this){
	    Console.WriteLine("Synchronizing table: "+tableName+". With command: "+tabDiffCmd);
	    new TableSynchronizer(conProps.getSourceServer(),conProps.getSourceDatabase(),tableName,conProps.getDestinationServer(),conProps.getDestinationDatabase(),tableName,tabDiffCmd );
		}
	      
   }
		public DatabaseSynchronizer(string sourceServer,  string sourceDB , string destinationServer,string destinationDB, string threadsStr){
			    conProps = new ConnectionProperty(sourceServer, destinationServer, sourceDB,destinationDB );
		        initConnections();
			    destinationTableList = getDestinationTables();
				int threads   =Int32.Parse(threadsStr);
			    syncThreads = new Thread[threads];
				int i =0;
				int counter =0;
			    foreach  (string tableName in destinationTableList){
				
						 if(!syncedTableList.Contains(tableName)){	
						 syncThreads[i]  = 	new Thread(() => synchTables("["+tableName+"]"));
								syncThreads[i].Start();
								syncedTableList.Add(tableName);
						 }
					i =checkRunningThreadCount();
				    while(i == threads){
						 Console.WriteLine("Waiting...");
						 Thread.Sleep(1000);
						 i =checkRunningThreadCount();
						 								  if(counter==destinationTableList.Count){
								 break;
								 
							 }
					}
					++counter;	
				}
			
			
	    }
		
				public DatabaseSynchronizer(string sourceServer,  string sourceDB , string destinationServer,string destinationDB,  string threadsStr, string option){
					string tabDiffCmd="";
					string[] tables;
					if(option.Contains("tabdiff.exe")){
			            tabDiffCmd = option;
						conProps = new ConnectionProperty(sourceServer, destinationServer, sourceDB,destinationDB );
						initConnections();
						destinationTableList = getDestinationTables();
						int threads   =Int32.Parse(threadsStr);
						syncThreads = new Thread[threads];
						int i =0;
						int counter =0;
						foreach  (string tableName in destinationTableList){
							syncThreads[i]  = 	new Thread(() => synchTables("["+tableName+"]",tabDiffCmd));
							if(!syncedTableList.Contains(tableName)){	
							syncThreads[i].Start();
							syncedTableList.Add(tableName);
							}
							i =checkRunningThreadCount();
							while(i == threads){
								 Console.WriteLine("Waiting...");
								 Thread.Sleep(1000);
								 i =checkRunningThreadCount();
								  if(counter==destinationTableList.Count){
								 break;
								 
							 }
							}
				++counter;
						}
			
				}else {
				
					tables = option.Split(',');
					Console.WriteLine("Synching tables: "+option);
					new DatabaseSynchronizer( sourceServer,   sourceDB ,  destinationServer, destinationDB,  threadsStr, tables);
					
					
				}
	    }
				public DatabaseSynchronizer(string sourceServer,  string sourceDB , string destinationServer,string destinationDB, string threadsStr,string[] tables){
					conProps = new ConnectionProperty(sourceServer, destinationServer, sourceDB,destinationDB );
					initConnections();
					string tableName = "";
					 if(tables.Length>0){
						int threads   =Int32.Parse(threadsStr);
						syncThreads = new Thread[threads];
						int i =0;
						int counter = 0;
					for (int j=0; j< tables.Length; j++){
	                  tableName = tables[j];
					  lock(this){
					     Console.WriteLine("Working on: "+tableName);
					  	syncThreads[i]  = 	new Thread(() => synchTables("["+tableName+"]"));
							if(!syncedTableList.Contains(tableName)){	
							while(syncThreads[i].IsAlive){
							 Console.WriteLine("Waiting...");
							 Thread.Sleep(1000);
							}
							syncThreads[i].Start();
							syncedTableList.Add(tableName);
							}			
						i =checkRunningThreadCount();
						while(i == threads){
							 Console.WriteLine("Waiting...");
							 Thread.Sleep(1000);
							 i =checkRunningThreadCount();
					 if(counter==destinationTableList.Count){
								 break;
								 
							 }
						}
					}
							++counter;					
						}
					}else{
						Console.WriteLine("There are no tables in the list.");
					}
			
	    }
		
		public DatabaseSynchronizer(string sourceServer,  string sourceDB , string destinationServer,string destinationDB, string threadsStr,string tableListStr, string taDiffCmd){
					conProps = new ConnectionProperty(sourceServer, destinationServer, sourceDB,destinationDB );
					initConnections();
					string[] tables = tableListStr.Split(',');
					if(tables.Length>0){
						destinationTableList.AddRange(tables);
						int threads   =Int32.Parse(threadsStr);
						syncThreads = new Thread[threads];
						int i =0;
						int counter = 0;
						foreach  (string tableName in destinationTableList){
				        syncThreads[i]  = 	new Thread(() => synchTables("["+tableName+"]",taDiffCmd));
						//if(!syncedTableList.Contains(tableName)){	
							syncThreads[i].Start();
							syncedTableList.Add(tableName);
					//	}
						 i =checkRunningThreadCount();
						while(i == threads){
							 Console.WriteLine("Waiting...");
							 Thread.Sleep(1000);
							 i =checkRunningThreadCount();
							 if(counter==destinationTableList.Count){
								 break;
								 
							 }
						}
									
						
						++counter;
						}
					}else{
						Console.WriteLine("There are no tables in the list.");
					}
			
	    }
		
		
	public void initConnections(){
            try
            {
                destinationConnection = new SqlConnection("Network Library=DBMSSOCN;Data Source=" + conProps.getDestinationServer() + ",1433;database=" +conProps.getDestinationDatabase()+ ";User id=" + conProps.getDestinationUser()+ ";Password=" +conProps.getDestinationPassword() + ";Connection Timeout=0;Pooling=false");
                destinationConnection.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error running table comparison: " + e.Message);
                Console.WriteLine(e.StackTrace);

            }

        }
		
		
		public string getSourceServer()
		   {
            return this.sourceServer;
        }
		
        public string getSourceDatabase()
        {
            return this.sourceDatabase;
        }
 

        public string getDestinationServer()
        {
            return this.destinationServer;
        }
        public string getDestinationDatabase()
        {
            return this.destinationDatabase;
        }
		
        public void setSourceServer(string server)
        {
            this.sourceServer = server;
        }

        public void setSourceDatabase(string database)
        {
            this.sourceDatabase = database;
        }

        public void setDestinationServer(string server)
        {
            this.destinationServer = server;
        }

      
        public void setDestinationDatabase(string database)
        {
            this.destinationDatabase = database;
        }
		
		public ArrayList getDestinationTables(){
			ArrayList tableList =  new ArrayList();
		  string sql_query = 
								";with fk_tables as (\n" +
								"	select	s1.name as from_schema\n" +
								"	,		o1.Name as from_table\n" +
								"	,		s2.name as to_schema\n" +
								"	,		o2.Name as to_table\n" +
								"	from	sys.foreign_keys fk\n" +
								"	inner	join sys.objects o1\n" +
								"	on		fk.parent_object_id = o1.object_id\n" +
								"	inner	join sys.schemas s1\n" +
								"	on		o1.schema_id = s1.schema_id\n" +
								"	inner	join sys.objects o2\n" +
								"	on		fk.referenced_object_id = o2.object_id\n" +
								"	inner	join sys.schemas s2\n" +
								"	on		o2.schema_id = s2.schema_id\n" +
								"	--For the purposes of finding dependency hierarchy \n" +
								"	--we're not worried about self-referencing tables\n" +
								"	where	not	(	s1.name = s2.name \n" +
								"				and	o1.name = o2.name)\n" +
								")\n" +
								",ordered_tables AS (\n" +
								"	SELECT	s.name as schemaName, t.name as tableName, 0 AS Level\n" +
								"	FROM	(	select	*\n" +
								"				from	sys.tables \n" +
								"				where	name <> 'sysdiagrams') t\n" +
								"	INNER	JOIN sys.schemas s\n" +
								"	on		t.schema_id = s.schema_id\n" +
								"	LEFT	OUTER JOIN fk_tables fk\n" +
								"	ON		s.name = fk.from_schema\n" +
								"	AND		t.name = fk.from_table\n" +
								"	WHERE	fk.from_schema IS NULL\n" +
								"	UNION	ALL\n" +
								"	SELECT	fk.from_schema, fk.from_table, ot.Level + 1\n" +
								"	FROM	fk_tables fk\n" +
								"	INNER	JOIN ordered_tables ot\n" +
								"	ON		fk.to_schema = ot.schemaName\n" +
								"	AND		fk.to_table = ot.tableName\n" +
								")\n" +
								"select	ot.tableName \n" +
								"from	ordered_tables ot\n" +
								"inner	join (\n" +
								"		select	schemaName,tableName,MAX(Level) maxLevel\n" +
								"		from	ordered_tables\n" +
								"		group	by schemaName,tableName\n" +
								") mx\n" +
								"on ot.schemaName = mx.schemaName\n" +
								"and ot.tableName = mx.tableName\n" +
								"and mx.maxLevel = ot.Level\n" +
								"ORDER	BY [Level] asc;\n"
								;
            
            try {

					Console.WriteLine("Running: "+sql_query);
                    SqlCommand cmd = new SqlCommand(sql_query, destinationConnection);
                    cmd.CommandTimeout = 0;
                    Console.WriteLine("Running SQL query: " + sql_query);
                    SqlDataReader dr = cmd.ExecuteReader();
					while (dr.HasRows)
        {
      
					 while(dr.Read()) 
                            {
                                tableList.Add(dr["tableName"].ToString().Trim());
                               Console.WriteLine("Succesffully added "+dr["tableName"].ToString().Trim() );
                            }
							 dr.NextResult();
		}
                
            } catch (Exception e)
            {
                Console.WriteLine("Error fetching tables: " + e.Message);
                Console.WriteLine(e.StackTrace);

            }	
			return tableList;
			
		}
				public static void Main (String[] args)
				{ try {
				             Console.WriteLine("Number of parameters: "+args.Length);
				             if(args.Length ==5){
						new DatabaseSynchronizer(args[0],args[1],args[2],args[3],args[4]);
						}else if(args.Length ==6){
							new DatabaseSynchronizer(args[0],args[1],args[2],args[3],args[4],args[5]);
						}else if(args.Length ==7){
							new DatabaseSynchronizer(args[0],args[1],args[2],args[3],args[4],args[5],args[6]);
						}

				} catch (Exception e)
            {
                Console.WriteLine("Error fetching tables: " + e.Message);
                Console.WriteLine(e.StackTrace);

            }	
				}
				
			



        }

    }
  

  
