/*
   Microsoft SQL Server Integration Services Script Task
   Write scripts using Microsoft Visual C# 2008.
   The ScriptMain is the entry point class of the script.
*/

using System;
using System.Data;
using System.IO;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;
using System.Data.SqlClient;
using System.Threading;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Diagnostics;


namespace Reco
{
    public partial class ReconcilationManager 
    {
			static string clientExcelConnectionStr = "";
			static string clientSettleConnectionStr = "";
			static Form prompt;
			static Thread mainThread;
			static string sqlConnectionStr = "";
			static string Query;
			static string outputPath= ".";
			static string outputFile= "";
			static  string serverName = "LOCALHOST";
			static  string excelFileLocation = ".";
			static string settleFileLocation = ".";
			static  string database = "postilion_office";
			static  string userName = "reportadmin";
			static  string password = "report.admin12";
			static  System.Data.DataTable dt = null;
			static  int INBUILT_QUERY_MODE = 1;
			static  int FILE_QUERY_MODE = 2;
			static int queryMode=INBUILT_QUERY_MODE;
			static SqlConnection officeConnection = null;
			static OleDbConnection excelConnection = null;
			static OleDbConnection settleConnection = null;
			static bool isConnected = false;
			static string reconQueryFile=Directory.GetCurrentDirectory()+"\\recon_data_check_query.sql";
			static string settleQueryFile=Directory.GetCurrentDirectory()+"\\settle_data_check_query.sql";
			public static string filters = "terminal_id,pan,retrieval_reference_nr,amount requested,sink_node_name";
			static bool continueRunning = true;
	
			static Label serverLabel = new Label() { Left = 10, Top = 15,Width =  70 , Height=20, Text = "Server: " };
			static TextBox serverTextBox = new TextBox() { Left =110, Top = 15, Width = 200 , Height=20};			
			static Label clientExcelLabel = new Label() { Left = 10, Top = 40 , Width = 100 , Height=20, Text = "Client Excel Path: " };
			static  TextBox clientExcelTextBox = new TextBox() { Left = 110, Top = 40, Width = 490 , Height=20};
			static Button clientBttn = new Button() { Text = "Browse", Left = 610, Width = 100, Top = 40};
			static Label settleExcelLabel = new Label() { Left = 10, Top = 65 , Width = 100 , Height=20, Text = "Settle Excel Path: " };
			static TextBox settleExcelTextBox = new TextBox() { Left = 110, Top = 65, Width = 490 , Height=20};
			static Button settleBttn = new Button() { Text = "Browse", Left = 610, Width = 100, Top = 65};
			static Label databaseLabel = new Label() { Left = 10, Top = 90 , Width = 70 , Height=20, Text = "Database: " };
			static TextBox databaseTextBox = new TextBox() { Left = 110, Top = 90, Width = 200 , Height=20};
			static Label usernameLabel = new Label() { Left = 10, Top = 120 , Width = 70 , Height=20, Text = "Username: " };
			static TextBox usernameTextBox = new TextBox() { Left = 110, Top = 120, Width = 200 , Height=20};
			static Label passwordLabel = new Label() { Left = 10, Top = 145 , Width = 70 , Height=20, Text = "Password: " };
			static TextBox passwordTextBox = new TextBox() { Left = 110, Top = 145, Width = 200 , Height=20,PasswordChar='*'};
			static Label outputLabel = new Label() { Left = 10, Top = 170 , Width = 70 , Height=20, Text = "Output File: " };
			static TextBox outputFileTextBox = new TextBox() { Left = 110, Top = 170, Width = 200 , Height=20};
			static Label outputPathLabel = new Label() { Left = 10, Top = 195 , Width =70 , Height=20, Text = "Output Path: " };
			static TextBox outputPathTextBox = new TextBox() { Left = 110, Top = 195, Width = 490 , Height=20};
			static Button outputBttn = new Button() { Text = "Browse", Left = 610, Width = 100, Top = 195};
			static TextBox displayBox = new TextBox() { Left = 110, Top = 220, Width = 600 , Height=250,Multiline = true,ScrollBars=ScrollBars.Vertical,ReadOnly =true};
			static  Button uploadBttn = new Button() { Text = "Upload", Left = 530, Width = 80, Top = 490 };
			static  Button MatchBttn = new Button() { Text = "Match", Left = 620, Width = 80, Top = 490}; // DialogResult = DialogResult.OK };
			static Label statusLabel = new Label() { Left = 10, Top = 515 , Width = 600 , Height=30 };
			const int SW_HIDE = 0;
			const int SW_SHOW = 5;
			static bool areFilesLoaded  =false;
			static MenuItem inBuilt ;
			static MenuItem sqlFile ;
			
			
			public  ReconcilationManager recoMan;

			[DllImport("kernel32.dll")]
			static extern IntPtr GetConsoleWindow();

			[DllImport("user32.dll")]
			static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        public  static  void getReconDetailsDialog(){
		   string caption = "Recon Tool"; 
             prompt = new Form()
            {
                Width = 790,
                Height = 600,
                FormBorderStyle = FormBorderStyle.Fixed3D,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
			 uploadBttn.Click+= new EventHandler(uploadFiles);
            MatchBttn.Click+= new EventHandler(startMatching);
		    prompt.Controls.Add(serverLabel);
            prompt.Controls.Add(serverTextBox);
		    prompt.Controls.Add(clientExcelLabel);
            prompt.Controls.Add(clientExcelTextBox);
			prompt.Controls.Add(settleExcelLabel);
            prompt.Controls.Add(settleExcelTextBox);
            prompt.Controls.Add(MatchBttn);
		    prompt.Controls.Add(uploadBttn);
			prompt.Controls.Add(clientBttn);
            prompt.Controls.Add(settleBttn);
			prompt.Controls.Add(databaseLabel);
            prompt.Controls.Add(databaseTextBox);
			prompt.Controls.Add(usernameLabel);
			prompt.Controls.Add(usernameTextBox);
			prompt.Controls.Add(passwordLabel);
			prompt.Controls.Add(passwordTextBox);
			prompt.Controls.Add(outputLabel);
			prompt.Controls.Add(outputFileTextBox);
			prompt.Controls.Add(outputPathLabel);
			prompt.Controls.Add(outputPathTextBox);
			prompt.Controls.Add(displayBox);
			prompt.Controls.Add(outputBttn);
			prompt.Controls.Add(statusLabel);
		
						
			prompt.Menu = new MainMenu();
			MenuItem item = new MenuItem("File"); 
			prompt.Menu.MenuItems.Add(item);
			item.MenuItems.Add("Clear Display", new EventHandler(clearDisplayBox));
            item.MenuItems.Add("Recon Query File", new EventHandler(getReconQuery));
            item.MenuItems.Add("Settle Query File", new EventHandler(getSettleQuery)); 
			item.MenuItems.Add("Show filters", new EventHandler(showSettltementFilter));
			MenuItem gItem = new MenuItem("Query Source");
			item.MenuItems.Add(gItem);
			item.MenuItems.Add("Exit", new EventHandler(closeRecon));
			 inBuilt = new MenuItem("In-built (Active)");
			 sqlFile = new MenuItem("SQL File");
			
			//gItem.MenuItems.Add("In-built", new EventHandler(setInBuiltQuerySource));
			//gItem.MenuItems.Add("SQL File", new EventHandler(setQueryFileSource));
			inBuilt.Click+= (sender, e) => { setInBuiltQuerySource(sender, e);};
			sqlFile.Click+= (sender, e) => { setQueryFileSource(sender, e);};
			gItem.MenuItems.Add(inBuilt);
			gItem.MenuItems.Add(sqlFile);
			

			MenuItem item2= new MenuItem("Info");
			prompt.Menu.MenuItems.Add(item2);
			item2.MenuItems.Add("Client Excel Format", new EventHandler(showClientExcelFormat));
			item2.MenuItems.Add("Settlement Excel Format", new EventHandler(showSettlementExcelFormat));
			item2.MenuItems.Add("Recon Query File Details", new EventHandler(showReconQueryFile));
			item2.MenuItems.Add("Settlement Query File Details", new EventHandler(showSettlementQueryFile));
			
						   
		clientBttn.Click+= (sender, e) => { 
	     OpenFileDialog openFileDialog1 = new OpenFileDialog();
				openFileDialog1.InitialDirectory = @".";
				openFileDialog1.Title = "Browse Excel Files";
				openFileDialog1.CheckFileExists = true;
				openFileDialog1.CheckPathExists = true;
				openFileDialog1.DefaultExt = "xls";
				openFileDialog1.Filter = "Excel files (*.xls*)|*.xls*";
				openFileDialog1.FilterIndex = 2;
				openFileDialog1.RestoreDirectory = true;
				openFileDialog1.ReadOnlyChecked = true;
				openFileDialog1.ShowReadOnly = true;
				if (openFileDialog1.ShowDialog() == DialogResult.OK)
				{
					clientExcelTextBox.Text = openFileDialog1.FileName;
				}
	   };
	   settleBttn.Click+= (sender, e) => { 
			OpenFileDialog openFileDialog2 = new OpenFileDialog();
			openFileDialog2.InitialDirectory = @".";
			openFileDialog2.Title = "Browse Excel Files";
			openFileDialog2.CheckFileExists = true;
			openFileDialog2.CheckPathExists = true;
			openFileDialog2.DefaultExt = "xls";
			openFileDialog2.Filter = "Excel files (*.xls*;)|*.xls*";
			openFileDialog2.FilterIndex = 2;
			openFileDialog2.RestoreDirectory = true;
			openFileDialog2.ReadOnlyChecked = true;
			openFileDialog2.ShowReadOnly = true;
			if (openFileDialog2.ShowDialog() == DialogResult.OK)
				{
					settleExcelTextBox.Text = openFileDialog2.FileName;
				}
	   };
		  outputBttn.Click+= (sender, e) => { 
	
		  
			OpenFileDialog openFileDialog3 = new OpenFileDialog();
			openFileDialog3.InitialDirectory = @".";
			openFileDialog3.Title = "Select Output Folder";
			openFileDialog3.ValidateNames = false;
		    openFileDialog3.Filter = "folders|*.neverseenthisfile";
			openFileDialog3.CheckFileExists = false;
			openFileDialog3.CheckPathExists = false;
		    openFileDialog3.FilterIndex = 2;
			openFileDialog3.RestoreDirectory = true;
			openFileDialog3.ReadOnlyChecked = true;
			openFileDialog3.ShowReadOnly = true;
			openFileDialog3.FileName = "Folder Selection";

			if (openFileDialog3.ShowDialog() == DialogResult.OK)
				{
					 outputPathTextBox.Text = openFileDialog3.FileName.Substring(0,openFileDialog3.FileName.Length -16);
				}
				
	   };
		    ThreadStart tsd = new ThreadStart(showPrompt);
				Thread trd = new Thread(tsd);
				trd.Start();
        }
		
public static void	showSettltementFilter(object sender, System.EventArgs e){
	
			   	    ThreadStart tsd = new ThreadStart(showFilterOptions);
				Thread formThread = new Thread(tsd);
				formThread.Start();

}
public static void  closeRecon(object sender, System.EventArgs e){
	 DialogResult dr = MessageBox.Show("Do you want to exit?", "Recon Tool - Exit Recon Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
		if(isConnected){
			settleConnection.Close();
			excelConnection.Close();
			officeConnection.Close();
			isConnected = false;
		}
		
        if(dr == DialogResult.Yes)
        {
           Environment.Exit(0);
        }
        
	
}

public static void  setInBuiltQuerySource(object sender, System.EventArgs e){
	 queryMode =  INBUILT_QUERY_MODE;
	 	 MenuItem sendObj = (MenuItem)sender;
		 sqlFile.Text = "SQL File";
		 inBuilt.Text="In-built (Active)";
	 MessageBox.Show("Query mode set to In-built queries.", "Recon Tool");
}
public static void  setQueryFileSource(object sender, System.EventArgs e){
	queryMode =  FILE_QUERY_MODE;
	 MenuItem sendObj = (MenuItem)sender;
	 sqlFile.Text = "SQL File (Active)";
	 inBuilt.Text="In-built";
	MessageBox.Show("Query mode set to External SQL files.", "Recon Tool");
}

public static void showPrompt(){
	prompt.ShowDialog();
	
}
public static void startMatching(object sender, System.EventArgs e){
		   	    ThreadStart tsd = new ThreadStart(runFileMatch);
				Thread minorThread = new Thread(tsd);
				minorThread.Start();
}


public static void uploadFiles(object sender, System.EventArgs e){
	if(!areFilesLoaded){
		    runFileUpload();
	} else{
		DialogResult dr = MessageBox.Show("Do you want to delete the data uploaded?", "Recon Tool - Reload tables", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if(dr == DialogResult.Yes)
        {
             runFileUpload();
        }
        

	}
			
			
}

public static void clearDisplayBox(object sender, System.EventArgs e){
	
	 displayBox.Text ="";
}
		public static void getReconQuery(object sender, System.EventArgs e){
		  OpenFileDialog openFileDialog1 = new OpenFileDialog();
				openFileDialog1.InitialDirectory = @".";
				openFileDialog1.Title = "Browse SQL Files";
				openFileDialog1.CheckFileExists = true;
				openFileDialog1.CheckPathExists = true;
				openFileDialog1.DefaultExt = "sql";
				openFileDialog1.Filter = "SQ files (*.sql)|*.sql";
				openFileDialog1.FilterIndex = 2;
				openFileDialog1.RestoreDirectory = true;
				openFileDialog1.ReadOnlyChecked = true;
				openFileDialog1.ShowReadOnly = true;
				if (openFileDialog1.ShowDialog() == DialogResult.OK)
				{
					reconQueryFile = openFileDialog1.FileName;
				}
				}

				
			public static void getSettleQuery(object sender, System.EventArgs e){
				OpenFileDialog openFileDialog1 = new OpenFileDialog();
				openFileDialog1.InitialDirectory = @".";
				openFileDialog1.Title = "Browse SQL Files";
				openFileDialog1.CheckFileExists = true;
				openFileDialog1.CheckPathExists = true;
				openFileDialog1.DefaultExt = "sql";
				openFileDialog1.Filter = "SQ files (*.sql)|*.sql";
				openFileDialog1.FilterIndex = 2;
				openFileDialog1.RestoreDirectory = true;
				openFileDialog1.ReadOnlyChecked = true;
				openFileDialog1.ShowReadOnly = true;
				if (openFileDialog1.ShowDialog() == DialogResult.OK)
				{
					settleQueryFile = openFileDialog1.FileName;
				}
				}
				
				public static void showClientExcelFormat(object sender, System.EventArgs e){
					
					MessageBox.Show("The Excel Files from  clients should have this format:\r\n[pan], [terminal_id], [card_acceptor_id_code], [merchant_type], [card_acceptor_name_loc], [message_type] ,[datetime_req], [system_trace_audit_nr],[retrieval_reference_nr],[auth_id_rsp]" ,"Recon Tool - Client Excel Format");
				}
				
				public static void showSettlementExcelFormat(object sender, System.EventArgs e){
					
					MessageBox.Show("The Excel Files for settlement should have this format:\r\n[[Host],[date],[F3],[card acceptor loc (Card acceptor id code)],[F5],[transaction date],[terminal id],[stan],[ptsp name],[pan],[transaction type],[account type],[response code description],[ptsp name1],[amount requested],[F16],[amount approved],[Retrieval reference number],[account number],[Ptsp fee],[merchant receivable],[merchant category],[F23],[transaction type 1],[transaction type 2],[F26]","Recon Tool - Settlement File Excel Format" );
				}
					public static void showReconQueryFile(object sender, System.EventArgs e){
					
					MessageBox.Show("FileName: "+reconQueryFile+"\r\n\r\n"+File.ReadAllText(reconQueryFile),"Recon Tool - Reconciliation Query File location" );
				}
					public static void showSettlementQueryFile(object sender, System.EventArgs e){
					
					MessageBox.Show( "FileName: "+settleQueryFile+"\r\n\r\n"+File.ReadAllText(settleQueryFile),"Recon Tool - Reconciliation Query File location" );
				}
        public  static  void initConnections()
        {
             try{
  
                serverName = serverTextBox.Text.Trim().Length!=0 ? serverTextBox.Text.Trim() : "LOCALHOST";
                excelFileLocation = clientExcelTextBox.Text.Trim();
				settleFileLocation = settleExcelTextBox.Text.Trim();
                database = databaseTextBox.Text.Trim().Length!=0? databaseTextBox.Text.Trim(): "postilion_office";
				outputFile = outputFileTextBox.Text.Trim().Length!=0?  outputFileTextBox.Text.Trim() : Directory.GetCurrentDirectory()+"\\recon_manager_output.csv";
                outputPath = outputPathTextBox.Text.Trim().Length!=0?  outputPathTextBox.Text.Trim() : "."; 
				outputFile = outputPath.Replace("\\","\\")+outputFile+".csv";
				outputFile = outputFile.Replace(".csv.csv",".csv");
				userName = usernameTextBox.Text.Trim().Length!=0? usernameTextBox.Text.Trim() : "reportadmin";
                password = passwordTextBox.Text.Trim().Length!=0 ? passwordTextBox.Text.Trim() : "report.admin12";
				displayBox.Text +="\r\n--------------------------------------------------";
				displayBox.Text +="\r\nRunning New Session ";
				displayBox.Text +="\r\n--------------------------------------------------";
                displayBox.Text +="\r\nSession Parameters:\r\n1.\tServerName: " + serverName + "\r\n2.\tDatabase: " + database + "\r\n3.\tExcel Source Path: " + excelFileLocation + "\r\n4.\tOutput File Name: " + outputFile + "\r\n5.\tUserName: " +userName + "\r\n6.\tPassword File: " + new String('*', password.Length)+ "\r\n7.\tSettlement Source File: " +settleFileLocation+ "\r\n7.\tOutput Folder: " +outputPath;
				displayBox.Text +="\r\n\r\n";
				
			if(excelFileLocation.Length !=0  && settleFileLocation.Length !=0  ){
            statusLabel.Text = "Connecting to "+ serverName + "...";
			if(isConnected){
					settleConnection.Close();
					excelConnection.Close();
					officeConnection.Close();
			}
            officeConnection = new SqlConnection("Network Library=DBMSSOCN;Data Source=" + serverName + ",1433;database=" + database + ";User id=reportadmin;Password=report.admin12;Connection Timeout=0");
            sqlConnectionStr = "Network Library=DBMSSOCN;Data Source=" + serverName + ",1433;database=" + database + ";User id="+userName+";Password="+password+";";
            officeConnection.Open();
            displayBox.Text +="\r\nSucessfully connected to " + serverName;
			statusLabel.Text = "Sucessfully connected to " + serverName;
          
            if (excelFileLocation.EndsWith("xls"))
            {
                excelFileLocation = "\"" + excelFileLocation + "\"";
                clientExcelConnectionStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES\"", excelFileLocation);
                // clientExcelConnectionStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFileLocation + ";Extended Properties=\"Excel 8.0;HDR=YES\"";
            } else if (excelFileLocation.EndsWith("xlsx")) {
                excelFileLocation = "\"" + excelFileLocation + "\"";
                clientExcelConnectionStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES\"", excelFileLocation);
                //clientExcelConnectionStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFileLocation + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
            }
            excelConnection = new OleDbConnection(clientExcelConnectionStr);
            try{
				excelConnection.Open();
			 }
			 catch(Exception e){
				     MessageBox.Show("Error opening Client Excel file\r\n"+e.Message, "Recon Tool - Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				    displayBox.Text +="\r\n"+e.Message;
				    displayBox.Text +="\r\n"+e.StackTrace;
				   displayBox.Text +="\r\nTrying another provider...";
				   displayBox.Text +=string.Format("\r\nProvider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES\"", excelFileLocation);
				   clientExcelConnectionStr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES\"", excelFileLocation);
				   excelConnection = new OleDbConnection(clientExcelConnectionStr);
				    try{
				   excelConnection.Open();
				   }
			 catch(Exception ex){
				   MessageBox.Show("Error opening Client Excel file\r\n"+ex.Message, "Recon Tool - Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				    displayBox.Text +="\r\n"+ex.Message;
				    displayBox.Text +="\r\n"+ex.StackTrace;
					continueRunning =false;
				   
				 
			 }
			 }
             displayBox.Text +="\r\nSucessfully opened Excel file  " + excelFileLocation;
			 statusLabel.Text = "Sucessfully opened Excel file  " + excelFileLocation;
			 
			 if(settleFileLocation.Length !=0  && settleFileLocation.Length !=0  ){
            if (settleFileLocation.EndsWith("xls"))
            {
                settleFileLocation = "\"" + settleFileLocation + "\"";
                clientSettleConnectionStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES\"", settleFileLocation);
                // clientSettleConnectionStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + settleFileLocation + ";Extended Properties=\"settle 8.0;HDR=YES\"";
            }
            else if (settleFileLocation.EndsWith("xlsx"))
            {
                settleFileLocation = "\"" + settleFileLocation + "\"";
                clientSettleConnectionStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES\"", settleFileLocation);
                //clientSettleConnectionStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + settleFileLocation + ";Extended Properties=\"settle 12.0 Xml;HDR=YES\"";
            }
            settleConnection = new OleDbConnection(clientSettleConnectionStr);
			
            try{
				settleConnection.Open();
			 }
			 catch(Exception e){
			 MessageBox.Show("Error opening Settlement Excel file\r\n"+e.Message, "Recon Tool - Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				    displayBox.Text +="\r\n"+e.Message;
				    displayBox.Text +="\r\n"+e.StackTrace;
				   displayBox.Text +="\r\nTrying another provider...";
				   displayBox.Text +=string.Format("\r\nProvider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"settle 8.0;HDR=YES\"", settleFileLocation);
				   clientSettleConnectionStr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"settle 8.0;HDR=YES\"", settleFileLocation);
				   settleConnection = new OleDbConnection(clientSettleConnectionStr);
			
				   
				    try{
					   settleConnection.Open();
				   }
			 catch(Exception ex){
				   MessageBox.Show("Error opening Settlement Excel file\r\n"+ex.Message, "Recon Tool - Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				    displayBox.Text +="\r\n"+ex.Message;
				    displayBox.Text +="\r\n"+ex.StackTrace;
					continueRunning =false;
			 }
				 
			 }
			  displayBox.Text +="\r\nSucessfully opened Excel file  " + settleFileLocation;
			  statusLabel.Text = "Sucessfully opened Excel file  "   + settleFileLocation;
			  isConnected = true;
             } else{
				 if(settleFileLocation.Length ==0  && settleFileLocation.Length ==0  ){
					  displayBox.Text +="\r\n Please provide the source path for the client Excel file and the Settlement file";
					  MessageBox.Show("Please provide the source path for the client Excel file  and the Settlement file", "Recon Tool");
					  statusLabel.Text = "There is no Excel file  to process..."; 
					 
				 }else  if(excelFileLocation.Length ==0){
					  displayBox.Text +="\r\n Please provide the source path for the client Excel file";
					  MessageBox.Show("Please provide the source path for the client Excel file", "Recon Tool");
					  statusLabel.Text = "There is no Excel file from the client to process...";
				  }else if(settleFileLocation.Length ==0){
					  displayBox.Text +="\r\n Please provide the source path for the Excel file for Settlement";
					  MessageBox.Show("Please provide the source path for the Excel file for Settlement", "Recon Tool");  
					  statusLabel.Text = "There is no Settlement file to compare to the client Excel file...";
				  }
				 
				 
			 }
			 
			
			 }
			 }
             catch (Exception e)
             {

				 displayBox.Text +="\r\n"+e.Message;
				 statusLabel.Text = e.Message;
				 displayBox.Text +="\r\n"+e.StackTrace;
                 MessageBox.Show("Error initiating connections\r\n"+e.Message, "Recon Tool - Critical Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				continueRunning =false;
 
			 }
			 
			 }
        static  void dropReconTables()
        {
				statusLabel.Text = "Dropping  reconciliation_data table...";
            try
            {

                SqlCommand thisCommand = officeConnection.CreateCommand();
                thisCommand.CommandText = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[recon_client_data_raw]') AND type in (N'U')) DROP TABLE [dbo].[recon_client_data_raw]";
                thisCommand.ExecuteNonQuery();
				thisCommand.CommandText = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[recon_client_data_settle_matched]') AND type in (N'U')) DROP TABLE [dbo].[recon_client_data_settle_matched]";
                thisCommand.ExecuteNonQuery();
			    thisCommand.CommandText = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[recon_client_data_settle_unmatched]') AND type in (N'U')) DROP TABLE [dbo].[recon_client_data_settle_unmatched]";
                thisCommand.ExecuteNonQuery();
				thisCommand.CommandText = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[recon_client_data_office_matched]') AND type in (N'U')) DROP TABLE [dbo].[recon_client_data_office_matched]";
                thisCommand.ExecuteNonQuery();
				thisCommand.CommandText = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[recon_client_data_office_unmatched]') AND type in (N'U')) DROP TABLE [dbo].[recon_client_data_office_unmatched]";
                thisCommand.ExecuteNonQuery();

            }
            catch (SqlException e)
            {
				 displayBox.Text +="\r\n"+e.Message;
				 displayBox.Text +="\r\n"+e.StackTrace;
				  MessageBox.Show("Error dropping reconciliation_data table \r\n"+e.Message, "Recon Tool - Error",MessageBoxButtons.OK, MessageBoxIcon.Warning);
				  continueRunning =false;
				   
            
            }
			statusLabel.Text = "Table has been removed";
        }

        static   void createReconTables(){

            try
            {

                SqlCommand thisCommand = officeConnection.CreateCommand();
                thisCommand.CommandTimeout = 0;
				
				string[] tableName = {"recon_client_data_raw", "recon_client_data_settle_matched", "recon_client_data_settle_unmatched", "recon_client_data_office_matched", "recon_client_data_office_unmatched"};
				for(int i=0; i< tableName.Length; i++){
				statusLabel.Text = "Creating "+tableName[i]+" table...";
				  displayBox.Text += "Creating "+tableName[i]+" table...";
                thisCommand.CommandText = "CREATE TABLE [dbo].["+tableName[i]+"]("
                                                    + "	[pan] [nvarchar](255) NULL,"
                                                    + "	[terminal_id] [nvarchar](255) NULL,"
                                                    + "	[card_acceptor_id_code] [nvarchar](255) NULL,"
                                                    + "	[merchant_type] [nvarchar](255) NULL,"
                                                    + "	[card_acceptor_name_loc] [nvarchar](255) NULL,"
                                                    + "	[message_type] [nvarchar](255) NULL,"
                                                    + "	[datetime_req] [nvarchar](255) NULL,"
                                                    + "	[system_trace_audit_nr] [nvarchar](255) NULL,"
                                                    + "	[retrieval_reference_nr] [nvarchar](255) NULL,"
                                                    + "	[auth_id_rsp] [nvarchar](255) NULL,"
                                                    + "	[F5] [nvarchar](255) NULL,"
													+ "	[amount requested] [float] NULL,"
                                                    + ") ON [PRIMARY]"
                                                    + "CREATE INDEX ix_pan  on ["+tableName[i]+"] ("
                                                    + "pan"
                                                    + ");"
                                                    + "CREATE INDEX ix_terminal_id  on ["+tableName[i]+"] ("
                                                    + "terminal_id"
                                                    + ");"
                                                    + "CREATE INDEX ix_card_acceptor_id_code  on ["+tableName[i]+"] ("
                                                    + "[card_acceptor_id_code]"
                                                    + ");"
                                                    + "CREATE INDEX ix_card_acceptor_name_loc on ["+tableName[i]+"] ("
                                                    + "[card_acceptor_name_loc]"
                                                    + ");"
                                                    + "CREATE INDEX ix_system_trace_audit_nr  on ["+tableName[i]+"] ("
                                                    + "[system_trace_audit_nr]"
                                                    + ");"
                                                    + "CREATE INDEX ix_retrieval_reference_nr  on ["+tableName[i]+"] ("
                                                    + "[system_trace_audit_nr]"
                                                    + ");"
                                                    + "CREATE INDEX ix_auth_id_rsp  on ["+tableName[i]+"] ("
                                                    + "auth_id_rsp"
                                                    + ");"
													+ "CREATE INDEX ix_amount_requested  on ["+tableName[i]+"] ("
                                                    + "[amount requested]"
                                                    + ");"
                                                    ;
                thisCommand.ExecuteNonQuery();
                displayBox.Text +="\r\nSucessfully created "+tableName[i]+" table ";
             }
            }
            catch (SqlException e)
            {
				 displayBox.Text +="\r\n"+e.Message;
				 displayBox.Text +="\r\n"+e.StackTrace;
				 MessageBox.Show("Error creating client reconciliation tables\r\n"+e.Message, "Recon Tool - Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				 continueRunning =false;
            }
        }
		
		  static  void createReconSettledTable()
        {

						try
						{
					    statusLabel.Text = "Creating recon_settled_transactions table for Settlement data...";
						SqlCommand thisCommand = officeConnection.CreateCommand();
						thisCommand.CommandTimeout = 0;
						thisCommand.CommandText = 
						"IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[recon_settled_transactions]') AND type in (N'U'))"
						+ "DROP TABLE [dbo].[recon_settled_transactions];"
						+ "CREATE TABLE [dbo].[recon_settled_transactions]("
						+ "	[Host] [nvarchar](255) NULL,"
						+ "	[date] [datetime] NULL,"
						+ "	[F3] [nvarchar](255) NULL,"
						+ "	[card acceptor loc (Card acceptor id code)] [nvarchar](255) NULL,"
						+ "	[F5] [nvarchar](255) NULL,"
						+ "	[transaction date] [datetime] NULL,"
						+ "	[terminal id] [nvarchar](255) NULL,"
						+ "	[stan] [float] NULL,"
						+ "	[ptsp name] [nvarchar](255) NULL,"
						+ "	[pan] [nvarchar](255) NULL,"
						+ "	[transaction type] [nvarchar](255) NULL,"
						+ "	[account type] [nvarchar](255) NULL,"
						+ "	[response code description] [nvarchar](255) NULL,"
						+ "	[ptsp name1] [nvarchar](255) NULL,"
						+ "	[amount requested] [float] NULL,"
						+ "	[F16] [float] NULL,"
						+ "	[amount approved] [float] NULL,"
						+ "	[Retrieval reference number] [float] NULL,"
						+ "	[account number] [float] NULL,"
						+ "	[Ptsp fee] [float] NULL,"
						+ "	[merchant receivable] [float] NULL,"
						+ "	[merchant category] [nvarchar](255) NULL,"
						+ "	[F23] [nvarchar](255) NULL,"
						+ "	[transaction type 1] [nvarchar](255) NULL,"
						+ "	[transaction type 2] [nvarchar](255) NULL,"
						+ "	[F26] [float] NULL"
						+ ") ON [PRIMARY];"
						+ "CREATE INDEX ix_tran_date ON  [recon_settled_transactions] ("
						+ "	[transaction date]"
						+ ")"
						+ "CREATE INDEX ix_term_id ON  [recon_settled_transactions] ("
						+ "	[terminal id]"
						+ ")"
						+ "CREATE INDEX ix_stan ON  [recon_settled_transactions] ("
						+ "	[stan]"
						+ ")"
						+ "CREATE INDEX ix_tran_type ON  [recon_settled_transactions] ("
						+ "	[transaction type]"
						+ ")"
						+ "CREATE INDEX ix_rrn ON  [recon_settled_transactions] ("
						+ "	[Retrieval reference number]"
						+ ")"
						+ "CREATE INDEX ix_rsp_code ON  [recon_settled_transactions] ("
						+ "	[response code description]"
						+ ")"
						+ "CREATE INDEX ix_amount_requested ON  [recon_settled_transactions] ("
						+ "	[amount requested]"
						+ ")";
						thisCommand.ExecuteNonQuery();
						displayBox.Text +="\r\nSucessfully created recon_settled_transactions table ";
						statusLabel.Text = "Sucessfully created recon_settled_transactions table.";

            }
            catch (SqlException e)
            {
				 displayBox.Text +="\r\n"+e.Message;
				 displayBox.Text +="\r\n"+e.StackTrace;
				statusLabel.Text = e.Message;
				MessageBox.Show("Error creating recon_settled_transactions table\r\n"+e.Message, "Recon Tool - Error",MessageBoxButtons.OK,MessageBoxIcon.Warning);
				continueRunning =false;
            }
        }

public  static  void InsertSettleExcelRecords()
        {

            try
            {
          
                    displayBox.Text +="\r\nReading  Settle Excel File: ";
					statusLabel.Text ="Reading  Settle Excel File: ";
                    dt = settleConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    String settleSheetName = "";
                    foreach (DataRow row in dt.Rows)
                    {
                        settleSheetName = row["TABLE_NAME"].ToString();
						displayBox.Text +="\r\nReading  Settle  sheet: "+settleSheetName;
						statusLabel.Text ="Reading  Settle  sheet: "+settleSheetName;
                        Query = string.Format("Select [Host],[date],[F3],[card acceptor loc (Card acceptor id code)],[F5],[transaction date],[terminal id],[stan],[ptsp name],[pan],[transaction type],[account type],[response code description],[ptsp name1],[amount requested],[F16],[amount approved],[Retrieval reference number],[account number],[Ptsp fee],[merchant receivable],[merchant category],[F23],[transaction type 1],[transaction type 2],[F26] FROM [{0}]", settleSheetName);
                  
						OleDbCommand Ecom = new OleDbCommand(Query, settleConnection);
						DataSet ds = new DataSet();
						OleDbDataAdapter oda = new OleDbDataAdapter(Query, settleConnection);
						displayBox.Text +="\r\nStarting Excel data upload for sheet: " + settleSheetName;
						statusLabel.Text ="Starting Excel data upload for sheet: " +settleSheetName; 
						oda.Fill(ds);  

						DataTable Exceldt = ds.Tables[0];

						SqlBulkCopy objbulk = new SqlBulkCopy(officeConnection);
						objbulk.BulkCopyTimeout = 0;
						//assigning Destination table name    

						objbulk.DestinationTableName = "recon_settled_transactions";
						//Mapping Table column    
						objbulk.ColumnMappings.Add("Host", "Host");
						objbulk.ColumnMappings.Add("date", "date");
						objbulk.ColumnMappings.Add("[card acceptor loc (Card acceptor id code)]", "[card acceptor loc (Card acceptor id code)]");
						objbulk.ColumnMappings.Add("F5", "F5");
						objbulk.ColumnMappings.Add("[transaction date]", "[transaction date]");
						objbulk.ColumnMappings.Add("[terminal id]", "[terminal id]");
						objbulk.ColumnMappings.Add("stan", "stan");
						objbulk.ColumnMappings.Add("[ptsp name]", "[ptsp name]");
						objbulk.ColumnMappings.Add("pan", "pan");
						objbulk.ColumnMappings.Add("[transaction type]", "[transaction type]");
						objbulk.ColumnMappings.Add("[account type]", "[account type]");
						objbulk.ColumnMappings.Add("[response code description]", "[response code description]");
						objbulk.ColumnMappings.Add("[ptsp name1]", "[ptsp name1]");
						objbulk.ColumnMappings.Add("[amount requested]", "[amount requested]");
						objbulk.ColumnMappings.Add("[F16]", "[F16]");
						objbulk.ColumnMappings.Add("[amount approved]", "[amount approved]");
						objbulk.ColumnMappings.Add("[Retrieval reference number]", "[Retrieval reference number]");
						objbulk.ColumnMappings.Add("[account number]", "[account number]");
						objbulk.ColumnMappings.Add("[Ptsp fee]", "[Ptsp fee]");
						objbulk.ColumnMappings.Add("[merchant receivable]", "[merchant receivable]");
						objbulk.ColumnMappings.Add("[merchant category]", "[merchant category]");
						objbulk.ColumnMappings.Add("[F23]", "[F23]");
						objbulk.ColumnMappings.Add("[transaction type 1]", "[transaction type 1]");
						objbulk.ColumnMappings.Add("[transaction type 2]", "[transaction type 2]");
						objbulk.ColumnMappings.Add("F26", "F26");
						 displayBox.Text +="\r\ninserting Datatable Records to DataBase...";
						objbulk.WriteToServer(Exceldt);                
						displayBox.Text +="\r\ninsert complete for Sheet: " + settleSheetName;
						statusLabel.Text ="insert complete for Sheet: " + settleSheetName;
                    }
                
         areFilesLoaded = true;
            }
            catch (Exception e)
            {

				 displayBox.Text +="\r\n"+e.Message;
				 displayBox.Text +="\r\n"+e.StackTrace;
				 statusLabel.Text=e.Message;
				  MessageBox.Show("Error importing Settlement Excel File\r\n"+e.Message, "Recon Tool - Error",MessageBoxButtons.OK,MessageBoxIcon.Warning);
				  continueRunning =false;
            
            }

        }
		
		   public static void matchSettledRecords() {
            try
            {
				 displayBox.Text +="\r\nMatching client information with Settlement data... ";
			    statusLabel.Text ="Matching client information with Settlement data...";
                string queryFile = Directory.GetCurrentDirectory()+"\\settle_data_check_query.sql";
				queryFile =  settleQueryFile.Length != 0? settleQueryFile: queryFile;				
				string sql_query="";
				if(queryMode==INBUILT_QUERY_MODE){
					string mainQry = "SET NOCOUNT ON; INSERT INTO [recon_client_data_settle_matched] SELECT   rec.[pan]"+ 
					",[terminal_id]"+ 
					",[card_acceptor_id_code]"+ 
					",[merchant_type]"+ 
					",[card_acceptor_name_loc]"+ 
					",[message_type]"+ 
					",[datetime_req]"+ 
					",[system_trace_audit_nr]"+ 
					",[retrieval_reference_nr]"+ 
					",[auth_id_rsp]"+ 
					",rec.[F5]"+ 
					",rec.[amount requested]"+ 
				"  FROM   [recon_client_data_raw] rec (NOLOCK) JOIN [recon_settled_transactions]  stl (NOLOCK)  ON ";
				string subQry = "";
				if(ReconcilationManager.filters.Contains("terminal_id"))  subQry  +="AND REPLICATE('0', 8-LEN( ltrim(rtrim(rec.[terminal_id]))))+ltrim(rtrim(rec.[terminal_id])) = REPLICATE('0', 8-LEN( ltrim(rtrim(stl.[terminal id]))))+ltrim(rtrim(stl.[terminal id]))"  ;
				if(ReconcilationManager.filters.Contains("pan"))subQry+= "AND LEFT(RTRIM(LTRIM(REPLACE(stl.[pan],' ', ''))),6) = LEFT(RTRIM(LTRIM(REPLACE(rec.pan,' ', ''))),6) AND  RIGHT(RTRIM(LTRIM(REPLACE(stl.[pan],' ', ''))),4) = RIGHT(RTRIM(LTRIM(REPLACE(rec.pan,' ', ''))),4)";
				if(ReconcilationManager.filters.Contains("retrieval_reference_nr"))subQry+="AND REPLICATE('0', 12-LEN( ltrim(rtrim(stl.[Retrieval reference number]))))+ltrim(rtrim(stl.[Retrieval reference number])) = REPLICATE('0', 12-LEN( ltrim(rtrim(rec.retrieval_reference_nr))))+ltrim(rtrim(rec.retrieval_reference_nr))"; 
				if(ReconcilationManager.filters.Contains("system_trace_audit_nr"))subQry+="AND REPLICATE('0', 6-LEN( ltrim(rtrim(stl.stan))))+ltrim(rtrim(stl.stan)) = REPLICATE('0', 6-LEN( ltrim(rtrim(rec.system_trace_audit_nr))))+ltrim(rtrim(rec.system_trace_audit_nr))";
				//if(ReconcilationManager.filters.Contains("system_trace_audit_nr"))subQry+="AND REPLICATE('0', 6-LEN( ltrim(rtrim(stl.stan))))+ltrim(rtrim(stl.stan)) = REPLICATE('0', 6-LEN( ltrim(rtrim(rec.system_trace_audit_nr))))+ltrim(rtrim(rec.system_trace_audit_nr))";
				if(ReconcilationManager.filters.Contains("amount requested"))subQry+="AND	stl.[amount requested] = rec.[amount requested] ";
				if(subQry.Length ==0)  subQry  += "AND REPLICATE('0', 12-LEN( ltrim(rtrim(stl.[Retrieval reference number]))))+ltrim(rtrim(stl.[Retrieval reference number])) = REPLICATE('0', 12-LEN( ltrim(rtrim(rec.retrieval_reference_nr))))+ltrim(rtrim(rec.retrieval_reference_nr))"; 
				subQry = subQry.Substring(4);
				
				sql_query =mainQry+subQry+" OPTION (RECOMPILE);";
				sql_query +=" INSERT INTO [recon_client_data_settle_unmatched]  SELECT * FROM [recon_client_data_raw] (NOLOCK) WHERE  [retrieval_reference_nr] NOT IN (  SELECT [retrieval_reference_nr] FROM [recon_client_data_settle_matched] (NOLOCK));  SELECT * FROM  [recon_settled_transactions] (NOLOCK) WHERE   REPLICATE('0', 12-LEN( ltrim(rtrim([Retrieval reference number]))))+ltrim(rtrim([Retrieval reference number]))  in (  SELECT  REPLICATE('0', 12-LEN( ltrim(rtrim(retrieval_reference_nr))))+ltrim(rtrim(retrieval_reference_nr)) FROM [recon_client_data_settle_matched] (NOLOCK))";
					
				}else if(queryMode==FILE_QUERY_MODE){		
				 displayBox.Text +="\r\nReading query file: " + queryFile;
				if (File.Exists(queryFile))
				{

					sql_query = File.ReadAllText(queryFile);
				}
				}
                SqlCommand cmd = new SqlCommand(sql_query, officeConnection);
                cmd.CommandTimeout = 0;
                displayBox.Text +="\r\nRunning SQL query: " + sql_query;
				statusLabel.Text ="Running SQL query to fetch settled transactions in client file";
                SqlDataReader dr = cmd.ExecuteReader();
				string outputFileLoc = outputFile.Substring(0,outputFile.LastIndexOf('.'))+"_client_settle_matched.csv";
				if(File.Exists(outputFileLoc))
					{
						File.Delete(outputFileLoc);
					}
                    displayBox.Text +="\r\nExporting results to  " + outputFileLoc;
					statusLabel.Text ="Exporting results to  " + outputFileLoc;

                    using (System.IO.StreamWriter fs = new System.IO.StreamWriter(outputFileLoc))
                    {

                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            string name = dr.GetName(i);
                            if (name.Contains(","))
                                name = "\"" + name + "\"";

                            fs.Write(name + ",");
                        }
                        fs.WriteLine();
                        while (dr.Read())
                        {
                            for (int i = 0; i < dr.FieldCount; i++)
                            {
                                string value = dr[i].ToString();
                                if (value.Contains(","))
                                    value = "\"" + value + "\"";

                                fs.Write(value + ",");
                            }
                            fs.WriteLine();
                        }


                        fs.Close();
						dr.Close();

                    }
                     displayBox.Text +="\r\nExport complete!" + outputFileLoc;
                     statusLabel.Text ="Export complete!" + outputFileLoc;
					 MessageBox.Show("Matching records with the settlement data have been successfully exported to:\r\n"+outputFileLoc, "Recon");
            }
            catch (Exception e)
            {
				  displayBox.Text +="\r\n"+e.Message;
				  displayBox.Text +="\r\n"+e.StackTrace;
				  statusLabel.Text =e.Message;
				  MessageBox.Show("Error exporting matched settlement records\r\n"+e.Message, "Recon Tool - Error",MessageBoxButtons.OK,MessageBoxIcon.Warning);
				  continueRunning =false;
            }
        }

        public  static  void InsertClientExcelRecords()
        {
					String excelSheetName = "";
            try
            {
          
                    displayBox.Text +="\r\nReading client Excel File: ";
					statusLabel.Text ="Reading client Excel File: ";
                    dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
					

                    // Add the sheet name to the string array.
                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheetName = row["TABLE_NAME"].ToString();

                        Query = string.Format("Select [pan], [terminal_id], [card_acceptor_id_code], [merchant_type], [card_acceptor_name_loc], [message_type] ,[datetime_req], [system_trace_audit_nr],[retrieval_reference_nr],[auth_id_rsp], [amount requested] FROM [{0}]", excelSheetName);
                  
                    OleDbCommand Ecom = new OleDbCommand(Query, excelConnection);
                    DataSet ds = new DataSet();
                    OleDbDataAdapter oda = new OleDbDataAdapter(Query, excelConnection);
					displayBox.Text +="\r\nStarting Excel data upload for sheet: " + excelSheetName;
					statusLabel.Text ="Starting Excel data upload for sheet: " + excelSheetName;
                    //creating object of SqlBulkCopy  
                    oda.Fill(ds);  

                    DataTable Exceldt = ds.Tables[0];

                    SqlBulkCopy objbulk = new SqlBulkCopy(officeConnection);
					objbulk.BulkCopyTimeout = 0;
                    //assigning Destination table name    

                    objbulk.DestinationTableName = "recon_client_data_raw";
                    //Mapping Table column    
                    objbulk.ColumnMappings.Add("pan", "pan");
                    objbulk.ColumnMappings.Add("terminal_id", "terminal_id");
                    objbulk.ColumnMappings.Add("card_acceptor_id_code", "card_acceptor_id_code");
                    objbulk.ColumnMappings.Add("merchant_type", "merchant_type");
                    objbulk.ColumnMappings.Add("card_acceptor_name_loc", "card_acceptor_name_loc");
                    objbulk.ColumnMappings.Add("message_type", "message_type");
                    objbulk.ColumnMappings.Add("datetime_req", "datetime_req");
                    objbulk.ColumnMappings.Add("system_trace_audit_nr", "system_trace_audit_nr");
                    objbulk.ColumnMappings.Add("retrieval_reference_nr", "retrieval_reference_nr");
                    objbulk.ColumnMappings.Add("auth_id_rsp", "auth_id_rsp");
					objbulk.ColumnMappings.Add("[amount requested]", "[amount requested]");
					
					
                    displayBox.Text +="\r\ninserting Datatable Records to DataBase...";
					statusLabel.Text ="inserting Datatable Records to DataBase...";
                    objbulk.WriteToServer(Exceldt);
                    
                    displayBox.Text +="\r\ninsert complete for Sheet: " + excelSheetName;
			        statusLabel.Text ="insert complete for Sheet: " + excelSheetName;
                    }
                
         
            }
            catch (Exception e)
            {
				 displayBox.Text +="\r\n"+e.Message;
				 displayBox.Text +="\r\n"+e.StackTrace;
				 statusLabel.Text =e.Message;
             MessageBox.Show("Error inserting data from "+ excelSheetName+" into reconciliation_data table: "+e.Message, "Recon Tool - Error",MessageBoxButtons.OK,MessageBoxIcon.Warning);
			 continueRunning =false;
            }

        }

        public static void matchClientRecords() {
			         
                    displayBox.Text +="\r\nMatching client records....";
					statusLabel.Text ="Matching client records....";
            try
            {
					
                string queryFile =  reconQueryFile!= ""? reconQueryFile: Directory.GetCurrentDirectory()+"\\recon_data_check_query.sql";
				string sql_query = "";
				if(queryMode==INBUILT_QUERY_MODE){
					string mainQry = "SET NOCOUNT ON; "+
					                 "IF(OBJECT_ID('tempdb.dbo.#TEMP_RECON_OFFICE') IS NOT NULL) DROP TABLE  #TEMP_RECON_OFFICE;"+
									" SELECT t.post_tran_cust_id ,abort_rsp_code ,acquirer_network_id ,payee ,pos_condition_code ,pos_entry_mode ,post_tran_id ,prev_post_tran_id ,prev_tran_approved ,pt_pos_card_input_mode , "+
									" realtime_business_date ,receiving_inst_id_code ,recon_business_date ,retention_data ,"+
									" t.retrieval_reference_nr ,routing_type ,rsp_code_req ,rsp_code_rsp ,settle_amount_impact ,settle_amount_req ,settle_amount_rsp ,settle_cash_req ,settle_cash_rsp ,settle_currency_code ,settle_entity_id ,"+
									" settle_proc_fee_req ,settle_proc_fee_rsp ,settle_tran_fee_req ,settle_tran_fee_rsp ,sink_node_name ,sponsor_bank,t.system_trace_audit_nr ,to_account_id ,to_account_type ,to_account_type_qualifier ,"+
									" tran_amount_req ,tran_amount_rsp ,tran_cash_req ,tran_cash_rsp ,tran_completed ,tran_currency_code ,tran_nr ,tran_postilion_originated ,tran_proc_fee_currency_code ,tran_proc_fee_req ,tran_proc_fee_rsp ,"+
									" tran_reversed ,tran_tran_fee_currency_code ,tran_tran_fee_req ,tran_tran_fee_rsp ,tran_type ,ucaf_data ,address_verification_data ,address_verification_result ,c.card_acceptor_id_code ,c.card_acceptor_name_loc ,"+
									" card_product ,card_seq_nr ,check_data ,draft_capture ,expiry_date ,mapped_card_acceptor_id_code ,c.merchant_type ,c.pan ,pan_encrypted ,pan_reference ,pan_search ,pos_card_capture_ability ,"+
									" pos_card_data_input_ability ,pos_card_data_input_mode ,pos_card_data_output_ability ,pos_card_present ,pos_cardholder_auth_ability ,pos_cardholder_auth_entity ,pos_cardholder_auth_method ,"+
									" pos_cardholder_present ,pos_operating_environment ,pos_pin_capture_ability ,pos_terminal_operator ,pos_terminal_output_ability ,pos_terminal_type ,service_restriction_code ,source_node_name ,c.terminal_id ,"+
									" terminal_owner ,totals_group ,acquiring_inst_id_code ,additional_rsp_data ,t.auth_id_rsp ,auth_reason ,auth_type ,bank_details ,batch_nr ,card_verification_result ,t.datetime_req ,datetime_rsp ,datetime_tran_gmt ,"+
									" datetime_tran_local ,extended_tran_type ,from_account_id ,from_account_type ,from_account_type_qualifier ,issuer_network_id ,message_reason_code ,t.message_type ,next_post_tran_id ,"+
									" online_system_id ,participant_id  INTO #TEMP_RECON_OFFICE FROM    post_tran   t(NOLOCK) JOIN post_tran_cust c ON t.post_tran_cust_id = c.post_tran_cust_id AND tran_postilion_originated = 0  JOIN recon_client_data_settle_unmatched rec (NOLOCK)   ON ";
					string subQry = "";
					if(ReconcilationManager.filters.Contains("terminal_id"))  subQry  +="AND REPLICATE('0', 8-LEN( ltrim(rtrim(rec.[terminal_id]))))+ltrim(rtrim(rec.[terminal_id]))  =   REPLICATE('0', 8-LEN( ltrim(rtrim(c.[terminal_id]))))+ltrim(rtrim(c.[terminal_id]))"  ;
					if(ReconcilationManager.filters.Contains("pan"))subQry+= "AND LEFT(RTRIM(LTRIM(REPLACE(rec.[pan],' ', ''))),6)=   LEFT(c.pan,6) AND  RIGHT(RTRIM(LTRIM(REPLACE(rec.[pan],' ', ''))),4) = RIGHT(c.pan,4)";
					if(ReconcilationManager.filters.Contains("retrieval_reference_nr"))subQry+=" AND REPLICATE('0', 12-LEN( ltrim(rtrim(rec.[retrieval_reference_nr]))))+ltrim(rtrim(rec.[retrieval_reference_nr])) =            REPLICATE('0', 12-LEN( ltrim(rtrim(t.retrieval_reference_nr))))+ltrim(rtrim(t.retrieval_reference_nr))"; 
					if(ReconcilationManager.filters.Contains("system_trace_audit_nr"))subQry+=" AND REPLICATE('0', 6-LEN( ltrim(rtrim(rec.stan))))+ltrim(rtrim(rec.stan)) = REPLICATE('0', 6-LEN( ltrim(rtrim(t.system_trace_audit_nr))))+ltrim(rtrim(t.system_trace_audit_nr))";
					if(ReconcilationManager.filters.Contains("sink_node_name"))subQry+=" AND CHARINDEX('SWT', t.sink_node_name) > 0 ";
					if(subQry.Length ==0)  subQry  += "AND REPLICATE('0', 12-LEN( ltrim(rtrim(rec.[retrieval_reference_nr]))))+ltrim(rtrim(rec.[retrieval_reference_nr])) = REPLICATE('0', 12-LEN( ltrim(rtrim(t.retrieval_reference_nr))))+ltrim(rtrim(t.retrieval_reference_nr))"; 
					subQry = subQry.Substring(4);
					sql_query =mainQry+subQry+" OPTION (RECOMPILE);";
					sql_query+= " ; INSERT INTO [recon_client_data_office_matched] SELECT    [pan]"+ 
											",[terminal_id]"+ 
											",[card_acceptor_id_code]"+ 
											",[merchant_type]"+ 
											",[card_acceptor_name_loc]"+ 
											",[message_type]"+ 
											",[datetime_req]"+ 
											",[system_trace_audit_nr]"+ 
											",[retrieval_reference_nr]"+ 
											",[auth_id_rsp]"+ 
											",[F5]"+ 
											",[amount requested] FROM  recon_client_data_settle_unmatched (NOLOCK) WHERE  REPLICATE('0', 12-LEN( ltrim(rtrim([retrieval_reference_nr]))))+ltrim(rtrim([retrieval_reference_nr])) IN (  SELECT  REPLICATE('0', 12-LEN( ltrim(rtrim([retrieval_reference_nr]))))+ltrim(rtrim([retrieval_reference_nr])) FROM #TEMP_RECON_OFFICE  )";
					sql_query +=" INSERT INTO [recon_client_data_office_unmatched]  SELECT * FROM [recon_client_data_raw] (NOLOCK) WHERE  REPLICATE('0', 12-LEN( ltrim(rtrim([retrieval_reference_nr]))))+ltrim(rtrim([retrieval_reference_nr])) NOT IN (  SELECT REPLICATE('0', 12-LEN( ltrim(rtrim([retrieval_reference_nr]))))+ltrim(rtrim([retrieval_reference_nr])) FROM [recon_client_data_office_matched] (NOLOCK));  SELECT   * FROM #TEMP_RECON_OFFICE (NOLOCK)";

					}else if(queryMode==FILE_QUERY_MODE){		
						  displayBox.Text +="\r\nReading query file: " + queryFile;
						  statusLabel.Text ="Reading query file: " + queryFile;
						if (File.Exists(queryFile))
						{

							sql_query = File.ReadAllText(queryFile);
						}
					}

                SqlCommand cmd = new SqlCommand(sql_query, officeConnection);
                cmd.CommandTimeout = 0;
                displayBox.Text +="\r\nRunning SQL query: " + sql_query;
				statusLabel.Text ="Running SQL query to fetch unsettled transactions that are not in Office";
                SqlDataReader dr = cmd.ExecuteReader();
                string outputFileLoc = outputFile.Substring(0,outputFile.LastIndexOf('.'))+"_unsettled_office_matched.csv";
				if(File.Exists(outputFileLoc))
					{
						File.Delete(outputFileLoc);
					}
                    displayBox.Text +="\r\nExporting results to  " + outputFileLoc;
					statusLabel.Text ="Exporting results to  " + outputFileLoc;

                    using (System.IO.StreamWriter fs = new System.IO.StreamWriter(outputFileLoc))
                    {

                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            string name = dr.GetName(i);
                            if (name.Contains(","))
                                name = "\"" + name + "\"";

                            fs.Write(name + ",");
                        }
                        fs.WriteLine();
                        while (dr.Read())
                        {
                            for (int i = 0; i < dr.FieldCount; i++)
                            {
                                string value = dr[i].ToString();
                                if (value.Contains(","))
                                    value = "\"" + value + "\"";

                                fs.Write(value + ",");
                            }
                            fs.WriteLine();
                        }


                        fs.Close();
					    dr.Close();
                    }
                    displayBox.Text +="\r\nExport complete!" + outputFileLoc;
					statusLabel.Text ="\nExport complete!" + outputFileLoc;
				     MessageBox.Show("Matching client records  from the server have been successfully exported to:\r\n"+outputFileLoc, "Recon");
            }
            catch (Exception e)
            {
				 displayBox.Text +="\r\n"+e.Message;
				 displayBox.Text +="\r\n"+e.StackTrace;
				 statusLabel.Text =e.Message;
				  MessageBox.Show("Error exporting matched client records: "+e.Message, "Recon Tool - Error",MessageBoxButtons.OK,MessageBoxIcon.Warning);
				  continueRunning =false;
            }
        }
		
		 public static void getUnmatchedClientRecords() {
			         
			displayBox.Text +="\r\n Exporting unmarched client records....";
			statusLabel.Text ="Exporting unmarched client client records....";
			string sql_query ="";
            try
            {
					        
                sql_query = "SELECT * FROM [recon_client_data_office_unmatched]  (NOLOCK)";
                SqlCommand cmd = new SqlCommand(sql_query, officeConnection);
                cmd.CommandTimeout = 0;
                displayBox.Text +="\r\nRunning SQL query: " + sql_query;
				statusLabel.Text ="Running SQL query to fetch transactions which were not settled and are not in Office";
                SqlDataReader dr = cmd.ExecuteReader();
                string outputFileLoc = outputFile.Substring(0,outputFile.LastIndexOf('.'))+"_unsettled_not_in_office.csv";
		if(File.Exists(outputFileLoc))
		{
			File.Delete(outputFileLoc);
		}
		displayBox.Text +="\r\nExporting results to  " + outputFileLoc;
		statusLabel.Text ="Exporting results to  " + outputFileLoc;

                    using (System.IO.StreamWriter fs = new System.IO.StreamWriter(outputFileLoc))
                    {

                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            string name = dr.GetName(i);
                            if (name.Contains(","))
                                name = "\"" + name + "\"";

                            fs.Write(name + ",");
                        }
                        fs.WriteLine();
                        while (dr.Read())
                        {
                            for (int i = 0; i < dr.FieldCount; i++)
                            {
                                string value = dr[i].ToString();
                                if (value.Contains(","))
                                    value = "\"" + value + "\"";

                                fs.Write(value + ",");
                            }
                            fs.WriteLine();
                        }


                        fs.Close();
					    dr.Close();
                    }
                    displayBox.Text +="\r\nExport complete!" + outputFileLoc;
					statusLabel.Text ="\nExport complete!" + outputFileLoc;
				     MessageBox.Show("Client records that could not be found in the  Office server have been successfully exported to:\r\n"+outputFileLoc, "Recon");
            }
            catch (Exception e)
            {
				 displayBox.Text +="\r\n"+e.Message;
				 displayBox.Text +="\r\n"+e.StackTrace;
				 statusLabel.Text =e.Message;
				  MessageBox.Show("Error exporting missing client records: "+e.Message, "Recon Tool - Error",MessageBoxButtons.OK,MessageBoxIcon.Warning);
				continueRunning =false;
		   }
        }
        public ReconcilationManager()
		
        {
			 
		   
        }

		public static void runFileUpload(){
			uploadBttn.Enabled =false;
			MatchBttn.Enabled =false;
			statusLabel.Text="Initiating connections..."; 
			initConnections();
	        if(excelFileLocation.Length !=0  && settleFileLocation.Length !=0 ){			
						 if(continueRunning ) dropReconTables();
						 if(continueRunning ) createReconSettledTable();
						 if(continueRunning ) createReconTables();
						 if(continueRunning ) InsertClientExcelRecords();
						 if(continueRunning ) createReconSettledTable();
						 if(continueRunning ) InsertSettleExcelRecords();
						
			 } else{
				  if(settleFileLocation.Length ==0  && settleFileLocation.Length ==0  ){
			  displayBox.Text +="\r\n Please provide the source path for the client Excel file and the Settlement file";
			  MessageBox.Show("Please provide the source path for the client Excel file  and the Settlement file", "Recon Tool");
			  statusLabel.Text = "There is no Excel file  to process..."; 
			         areFilesLoaded = false;
		 }else  if(excelFileLocation.Length ==0){
			  displayBox.Text +="\r\n Please provide the source path for the client Excel file";
			  MessageBox.Show("Please provide the source path for the client Excel file", "Recon Tool");
			  statusLabel.Text = "There is no Excel file from the client to process...";
			  areFilesLoaded = false;
		  }else if(settleFileLocation.Length ==0){
			  displayBox.Text +="\r\n Please provide the source path for the Excel file for Settlement";
			  MessageBox.Show("Please provide the source path for the Excel file for Settlement", "Recon Tool");  
			  statusLabel.Text = "There is no Settlement file to compare to the client Excel file...";
			  areFilesLoaded = false;
		  }
			  
			 }
			  
             if(areFilesLoaded){ 
				MessageBox.Show("Client and Settlement data have been successfully loaded .", "Recon Tool");
				uploadBttn.Text = "Reload";
			 } 
			  uploadBttn.Enabled =true;
			  MatchBttn.Enabled =true;
			  continueRunning = true;
		}
		 
		 public static void  runFileMatch(){
			  if(areFilesLoaded){
				    MatchBttn.Enabled =false;
					uploadBttn.Enabled = false;
				 if(continueRunning ) 	matchSettledRecords();
				 if(continueRunning ) 	matchClientRecords();
				 if(continueRunning ) 	getUnmatchedClientRecords();
			 	    prompt.DialogResult = DialogResult.None; 
				    MatchBttn.Enabled =true;
					uploadBttn.Enabled = true;
				 if(continueRunning ) 	MessageBox.Show("All files have been successfully exported .", "Recon Tool");
			  } else {
				  MessageBox.Show("Please upload client and settlement files first.", "Recon Tool");  
			  }
			  continueRunning = true;
		 }
		 
		 
		  public static void showFilterOptions(){
		//	System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);  
			System.Windows.Forms.Application.EnableVisualStyles();
			Form1 f = new Form1();
			System.Windows.Forms.Application.Run(f);
			
		}
    public static void Main(String[] args){
				var handle = GetConsoleWindow();
				ShowWindow(handle, SW_HIDE);
		   	    ThreadStart tsd = new ThreadStart(getReconDetailsDialog);
				mainThread = new Thread(tsd);
				mainThread.Start();
      }
	}

    

  
  
  partial class Form1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && ( components != null )) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
		
		void clossDialog(object sender, System.EventArgs e){
			ReconcilationManager.filters  = ccb.getAllCheckedItems();
			Dispose(true);
			
		}
        private void InitializeComponent() {
            this.txtOut = new System.Windows.Forms.TextBox();
			closeBttn = new Button() { Text = "OK", Left = 320, Width = 100, Top = 80};
			closeBttn.Click+= new EventHandler(clossDialog);
            this.ccb = new CheckedComboBox();
            this.SuspendLayout();
            // 
            // txtOut
            // 
            this.txtOut.Location = new System.Drawing.Point(12, 162);
            this.txtOut.Multiline = true;
		
            this.txtOut.Name = "txtOut";
            this.txtOut.Size = new System.Drawing.Size(400, 132);
            this.txtOut.TabIndex = 1;
	
            // 
            // ccb
            // 
            this.ccb.CheckOnClick = true;
            this.ccb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.ccb.DropDownHeight = 1;
            this.ccb.FormattingEnabled = true;
            this.ccb.IntegralHeight = false;
            this.ccb.Location = new System.Drawing.Point(12, 22);
            this.ccb.Name = "ccb";
            this.ccb.Size = new System.Drawing.Size(400, 21);
            this.ccb.TabIndex = 0;
            this.ccb.ValueSeparator = ", ";
            this.ccb.DropDownClosed += new System.EventHandler(this.ccb_DropDownClosed);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(430, 306);
            this.Controls.Add(this.txtOut);
			this.Controls.Add(this.closeBttn);
            this.Controls.Add(this.ccb);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Edit Query Filter";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private CheckedComboBox ccb;
        private System.Windows.Forms.TextBox txtOut;
		private Button closeBttn;
    }
    
        public partial class Form1 : Form {
			
			
            private string[] queryOptsArr = { "terminal_id", "pan", "retrieval_reference_nr", "system_trace_audit_nr", "amount requested","sink_node_name" };
            // ,"A very long string exceeding the dropdown width and forcing a scrollbar to appear to make the content viewable"};
    
            public Form1() {
                InitializeComponent();
                // Manually add handler for when an item check state has been modified.
                ccb.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.ccb_ItemCheck);
            }
    
            private void Form1_Load(object sender, EventArgs e) {
                for (int i = 0; i < queryOptsArr.Length; i++) {
                    CCBoxItem item = new CCBoxItem(queryOptsArr[i], i);
                    ccb.Items.Add(item);
                }
                // If more then 5 items, add a scroll bar to the dropdown.
                ccb.MaxDropDownItems = 5;
                // Make the "Name" property the one to display, rather than the ToString() representation.
                ccb.DisplayMember = "Name";
                ccb.ValueSeparator = ", ";
                // Check the first 2 items.
				
				  if(ReconcilationManager.filters.Contains("terminal_id"))  ccb.SetItemChecked(0, true);
                  if(ReconcilationManager.filters.Contains("pan")) ccb.SetItemChecked(1, true);
				  if(ReconcilationManager.filters.Contains("retrieval_reference_nr"))ccb.SetItemChecked(2, true);
				  if(ReconcilationManager.filters.Contains("system_trace_audit_nr"))ccb.SetItemChecked(3, true);
				  if(ReconcilationManager.filters.Contains("amount requested"))ccb.SetItemChecked(4, true);
				   if(ReconcilationManager.filters.Contains("amount requested"))ccb.SetItemChecked(5, true);
			  
                //ccb.SetItemCheckState(1, CheckState.Indeterminate);
            }
    
            private void ccb_DropDownClosed(object sender, EventArgs e) {
               // txtOut.AppendText("DropdownClosed\r\n");
                txtOut.AppendText(string.Format("value changed: {0}\r\n", ccb.ValueChanged));
                txtOut.AppendText(string.Format("value: {0}\r\n", ccb.Text));
                // Display all checked items.
                StringBuilder sb = new StringBuilder("filters Selected: ");
                foreach (CCBoxItem item in ccb.CheckedItems) {
                    sb.Append(item.Name).Append(ccb.ValueSeparator);
                }
                sb.Remove(sb.Length-ccb.ValueSeparator.Length, ccb.ValueSeparator.Length);
                txtOut.AppendText(sb.ToString());
                txtOut.AppendText("\r\n");
            }
    
            private void ccb_ItemCheck(object sender, ItemCheckEventArgs e) {
                CCBoxItem item = ccb.Items[e.Index] as CCBoxItem;
				 bool itemCheckedState = true;
		
				  foreach (CCBoxItem item1 in ccb.CheckedItems) {
                     if(item.Name ==item1.Name ) {
						itemCheckedState = false; 
						break; 
					 }
                }
				
				if (itemCheckedState){
                    txtOut.AppendText(string.Format(" '{0}'  has  been included in the query filter\r\n", item.Name));
				} else {
				   txtOut.AppendText(string.Format(" '{0}'  has  been removed from the query filter\r\n", item.Name));
					
				}
				
            }   
				
    }
	
	  public class CCBoxItem {
        private int val;
        public int Value {
            get { return val; }
            set { val = value; }
        }
        
        private string name;
        public string Name {
            get { return name; }
            set { name = value; }
        }

        public CCBoxItem() {
        }

        public CCBoxItem(string name, int val) {
            this.name = name;
            this.val = val;
        }

        public override string ToString() {
            return string.Format("name: '{0}', value: {1}", name, val);
        }
    }
	
	  public class CheckedComboBox : ComboBox {
        /// <summary>
        /// Internal class to represent the dropdown list of the CheckedComboBox
        /// </summary>
        internal class Dropdown : Form {
            // ---------------------------------- internal class CCBoxEventArgs --------------------------------------------
            /// <summary>
            /// Custom EventArgs encapsulating value as to whether the combo box value(s) should be assignd to or not.
            /// </summary>
            internal class CCBoxEventArgs : EventArgs {
                private bool assignValues;
                public bool AssignValues {
                    get { return assignValues; }
                    set { assignValues = value; }
                }
                private EventArgs e;
                public EventArgs EventArgs {
                    get { return e; }
                    set { e = value; }
                }
                public CCBoxEventArgs(EventArgs e, bool assignValues) : base() {
                    this.e = e;
                    this.assignValues = assignValues;
                }
            }

            // ---------------------------------- internal class CustomCheckedListBox --------------------------------------------

            /// <summary>
            /// A custom CheckedListBox being shown within the dropdown form representing the dropdown list of the CheckedComboBox.
            /// </summary>
            internal class CustomCheckedListBox : CheckedListBox {
                private int curSelIndex = -1;

                public CustomCheckedListBox() : base() {
                    this.SelectionMode = SelectionMode.One;
                    this.HorizontalScrollbar = true;                    
                }

                /// <summary>
                /// Intercepts the keyboard input, [Enter] confirms a selection and [Esc] cancels it.
                /// </summary>
                /// <param name="e">The Key event arguments</param>
                protected override void OnKeyDown(KeyEventArgs e) {
                    if (e.KeyCode == Keys.Enter) {
                        // Enact selection.
                        ((CheckedComboBox.Dropdown) Parent).OnDeactivate(new CCBoxEventArgs(null, true));
                        e.Handled = true;

                    } else if (e.KeyCode == Keys.Escape) {
                        // Cancel selection.
                        ((CheckedComboBox.Dropdown) Parent).OnDeactivate(new CCBoxEventArgs(null, false));
                        e.Handled = true;

                    } else if (e.KeyCode == Keys.Delete) {
                        // Delete unckecks all, [Shift + Delete] checks all.
                        for (int i = 0; i < Items.Count; i++) {
                            SetItemChecked(i, e.Shift);
                        }
                        e.Handled = true;
                    }
                    // If no Enter or Esc keys presses, let the base class handle it.
                    base.OnKeyDown(e);
                }

                protected override void OnMouseMove(MouseEventArgs e) {
                    base.OnMouseMove(e);
                    int index = IndexFromPoint(e.Location);
                    Debug.WriteLine("Mouse over item: " + (index >= 0 ? GetItemText(Items[index]) : "None"));
                    if ((index >= 0) && (index != curSelIndex)) {
                        curSelIndex = index;
                        SetSelected(index, true);
                    }
                }

            } // end internal class CustomCheckedListBox

            // --------------------------------------------------------------------------------------------------------

            // ********************************************* Data *********************************************

            private CheckedComboBox ccbParent;

            // Keeps track of whether checked item(s) changed, hence the value of the CheckedComboBox as a whole changed.
            // This is simply done via maintaining the old string-representation of the value(s) and the new one and comparing them!
            private string oldStrValue = "";
            public bool ValueChanged {
                get {
                    string newStrValue = ccbParent.Text;
                    if ((oldStrValue.Length > 0) && (newStrValue.Length > 0)) {
                        return (oldStrValue.CompareTo(newStrValue) != 0);
                    } else {
                        return (oldStrValue.Length != newStrValue.Length);
                    }
                }
            }

            // Array holding the checked states of the items. This will be used to reverse any changes if user cancels selection.
            bool[] checkedStateArr;

            // Whether the dropdown is closed.
            private bool dropdownClosed = true;

            private CustomCheckedListBox cclb;
            public CustomCheckedListBox List {
                get { return cclb; }
                set { cclb = value; }
            }

            // ********************************************* Construction *********************************************

            public Dropdown(CheckedComboBox ccbParent) {
                this.ccbParent = ccbParent;
                InitializeComponent();
                this.ShowInTaskbar = false;
                // Add a handler to notify our parent of ItemCheck events.
                this.cclb.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.cclb_ItemCheck);
            }

            // ********************************************* Methods *********************************************

            // Create a CustomCheckedListBox which fills up the entire form area.
            private void InitializeComponent() {
                this.cclb = new CustomCheckedListBox();
                this.SuspendLayout();
                // 
                // cclb
                // 
                this.cclb.BorderStyle = System.Windows.Forms.BorderStyle.None;
                this.cclb.Dock = System.Windows.Forms.DockStyle.Fill;
                this.cclb.FormattingEnabled = true;
				Rectangle rect = RectangleToScreen(this.ClientRectangle);
                this.cclb.Location =  new Point(rect.X, rect.Y + this.Size.Height);
                this.cclb.Name = "cclb";
                this.cclb.Size = new System.Drawing.Size(47, 15);
                this.cclb.TabIndex = 0;
                // 
                // Dropdown
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.BackColor = System.Drawing.SystemColors.Menu;
                this.ClientSize = new System.Drawing.Size(47, 16);
                this.ControlBox = false;
                this.Controls.Add(this.cclb);
                this.ForeColor = System.Drawing.SystemColors.ControlText;
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
                this.MinimizeBox = false;
                this.Name = "ccbParent";
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.ResumeLayout(false);
            }

            public string GetCheckedItemsStringValue() {
                StringBuilder sb = new StringBuilder("");
                for (int i = 0; i < cclb.CheckedItems.Count; i++) {                    
                    sb.Append(cclb.GetItemText(cclb.CheckedItems[i])).Append(ccbParent.ValueSeparator);
                }
                if (sb.Length > 0) {
                    sb.Remove(sb.Length - ccbParent.ValueSeparator.Length, ccbParent.ValueSeparator.Length);
                }
                return sb.ToString();
            }

            /// <summary>
            /// Closes the dropdown portion and enacts any changes according to the specified boolean parameter.
            /// NOTE: even though the caller might ask for changes to be enacted, this doesn't necessarily mean
            ///       that any changes have occurred as such. Caller should check the ValueChanged property of the
            ///       CheckedComboBox (after the dropdown has closed) to determine any actual value changes.
            /// </summary>
            /// <param name="enactChanges"></param>
            public void CloseDropdown(bool enactChanges) {
                if (dropdownClosed) {
                  return;
                }                
                Debug.WriteLine("CloseDropdown");
                // Perform the actual selection and display of checked items.
                if (enactChanges) {
                    ccbParent.SelectedIndex = -1;                    
                    // Set the text portion equal to the string comprising all checked items (if any, otherwise empty!).
                    ccbParent.Text = GetCheckedItemsStringValue();

                } else {
                    // Caller cancelled selection - need to restore the checked items to their original state.
                    for (int i = 0; i < cclb.Items.Count; i++) {
                        cclb.SetItemChecked(i, checkedStateArr[i]);
                    }
                }
                // From now on the dropdown is considered closed. We set the flag here to prevent OnDeactivate() calling
                // this method once again after hiding this window.
                dropdownClosed = true;
                // Set the focus to our parent CheckedComboBox and hide the dropdown check list.
                ccbParent.Focus();
                this.Hide();
                // Notify CheckedComboBox that its dropdown is closed. (NOTE: it does not matter which parameters we pass to
                // OnDropDownClosed() as long as the argument is CCBoxEventArgs so that the method knows the notification has
                // come from our code and not from the framework).
                ccbParent.OnDropDownClosed(new CCBoxEventArgs(null, false));
            }

            protected override void OnActivated(EventArgs e) {
                Debug.WriteLine("OnActivated");
                base.OnActivated(e);
                dropdownClosed = false;
                // Assign the old string value to compare with the new value for any changes.
                oldStrValue = ccbParent.Text;
                // Make a copy of the checked state of each item, in cace caller cancels selection.
                checkedStateArr = new bool[cclb.Items.Count];
                for (int i = 0; i < cclb.Items.Count; i++) {
                    checkedStateArr[i] = cclb.GetItemChecked(i);
                }
            }

            protected override void OnDeactivate(EventArgs e) {
                Debug.WriteLine("OnDeactivate");
                base.OnDeactivate(e);
                CCBoxEventArgs ce = e as CCBoxEventArgs;
                if (ce != null) {
                    CloseDropdown(ce.AssignValues);

                } else {
                    // If not custom event arguments passed, means that this method was called from the
                    // framework. We assume that the checked values should be registered regardless.
                    CloseDropdown(true);
                }
            }

            private void cclb_ItemCheck(object sender, ItemCheckEventArgs e) {
                if (ccbParent.ItemCheck != null) {
                    ccbParent.ItemCheck(sender, e);
                }
            }

        } // end internal class Dropdown

        // ******************************** Data ********************************
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        // A form-derived object representing the drop-down list of the checked combo box.
        private Dropdown dropdown;

        // The valueSeparator character(s) between the ticked elements as they appear in the 
        // text portion of the CheckedComboBox.
        private string valueSeparator;
        public string ValueSeparator {
            get { return valueSeparator; }
            set { valueSeparator = value; }
        }

        public bool CheckOnClick {
            get { return dropdown.List.CheckOnClick; }
            set { dropdown.List.CheckOnClick = value; }
        }

        public new string DisplayMember {
            get { return dropdown.List.DisplayMember; }
            set { dropdown.List.DisplayMember = value; }
        }

        public new CheckedListBox.ObjectCollection Items {
            get { return dropdown.List.Items; }
        }

        public CheckedListBox.CheckedItemCollection CheckedItems {
            get { return dropdown.List.CheckedItems; }
        }
        
        public CheckedListBox.CheckedIndexCollection CheckedIndices {
            get { return dropdown.List.CheckedIndices; }
        }

        public bool ValueChanged {
            get { return dropdown.ValueChanged; }
        }

        // Event handler for when an item check state changes.
        public event ItemCheckEventHandler ItemCheck;
        
        // ******************************** Construction ********************************

        public CheckedComboBox() : base() {
            // We want to do the drawing of the dropdown.
            this.DrawMode = DrawMode.OwnerDrawVariable;
            // Default value separator.
            this.valueSeparator = ", ";
            // This prevents the actual ComboBox dropdown to show, although it's not strickly-speaking necessary.
            // But including this remove a slight flickering just before our dropdown appears (which is caused by
            // the empty-dropdown list of the ComboBox which is displayed for fractions of a second).
            this.DropDownHeight = 1;            
            // This is the default setting - text portion is editable and user must click the arrow button
            // to see the list portion. Although we don't want to allow the user to edit the text portion
            // the DropDownList style is not being used because for some reason it wouldn't allow the text
            // portion to be programmatically set. Hence we set it as editable but disable keyboard input (see below).
            this.DropDownStyle = ComboBoxStyle.DropDown;
            this.dropdown = new Dropdown(this);
            // CheckOnClick style for the dropdown (NOTE: must be set after dropdown is created).
            this.CheckOnClick = true;
        }

        // ******************************** Operations ********************************

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }        

        protected override void OnDropDown(EventArgs e) {
            base.OnDropDown(e);
            DoDropDown();    
        }

        private void DoDropDown() {
            if (!dropdown.Visible) {
                Rectangle rect = RectangleToScreen(this.ClientRectangle);
                dropdown.Location = new Point(rect.X, rect.Y + this.Size.Height);
                int count = dropdown.List.Items.Count;
                if (count > this.MaxDropDownItems) {
                    count = this.MaxDropDownItems;
                } else if (count == 0) {
                    count = 1;
                }
                dropdown.Size = new Size(this.Size.Width, (dropdown.List.ItemHeight) * count + 2);
                dropdown.Show(this);
            }
        }

        protected override void OnDropDownClosed(EventArgs e) {
            // Call the handlers for this event only if the call comes from our code - NOT the framework's!
            // NOTE: that is because the events were being fired in a wrong order, due to the actual dropdown list
            //       of the ComboBox which lies underneath our dropdown and gets involved every time.
            if (e is Dropdown.CCBoxEventArgs) {
                base.OnDropDownClosed(e);
            }
        }

        protected override void OnKeyDown(KeyEventArgs e) {
            if (e.KeyCode == Keys.Down) {
                // Signal that the dropdown is "down". This is required so that the behaviour of the dropdown is the same
                // when it is a result of user pressing the Down_Arrow (which we handle and the framework wouldn't know that
                // the list portion is down unless we tell it so).
                // NOTE: all that so the DropDownClosed event fires correctly!                
                OnDropDown(null);
            }
            // Make sure that certain keys or combinations are not blocked.
            e.Handled = !e.Alt && !(e.KeyCode == Keys.Tab) &&
                !((e.KeyCode == Keys.Left) || (e.KeyCode == Keys.Right) || (e.KeyCode == Keys.Home) || (e.KeyCode == Keys.End));

            base.OnKeyDown(e);
        }

        protected override void OnKeyPress(KeyPressEventArgs e) {
            e.Handled = true;
            base.OnKeyPress(e);
        }

        public bool GetItemChecked(int index) {
            if (index < 0 || index > Items.Count) {
                throw new ArgumentOutOfRangeException("index", "value out of range");
            } else {
                return dropdown.List.GetItemChecked(index);
            }
        }

        public void SetItemChecked(int index, bool isChecked) {
            if (index < 0 || index > Items.Count) {
                throw new ArgumentOutOfRangeException("index", "value out of range");
            } else {
                dropdown.List.SetItemChecked(index, isChecked);
                // Need to update the Text.
                this.Text = dropdown.GetCheckedItemsStringValue();
            }
        }

        public CheckState GetItemCheckState(int index) {
            if (index < 0 || index > Items.Count) {
                throw new ArgumentOutOfRangeException("index", "value out of range");
            } else {
                return dropdown.List.GetItemCheckState(index);
            }
        }

        public void SetItemCheckState(int index, CheckState state) {
            if (index < 0 || index > Items.Count) {
                throw new ArgumentOutOfRangeException("index", "value out of range");
            } else {
                dropdown.List.SetItemCheckState(index, state);
                // Need to update the Text.
                this.Text = dropdown.GetCheckedItemsStringValue();
            }
        }
		public string  getAllCheckedItems (){
			
			return dropdown.GetCheckedItemsStringValue();
		}
    }
}