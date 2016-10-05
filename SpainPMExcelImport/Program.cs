/*
 * Created by SharpDevelop.
 * User: val01039
 * Date: 4.10.2016
 * Time: 13:12
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using PowerMILL;
using System.IO;

namespace SpainPMExcelImport
{
	class Program
	{
		public static void Main(string[] args)
		{
			Application pmApp = null;
			
			try {
				pmApp = (PowerMILL.Application) System.Runtime.InteropServices.Marshal.GetActiveObject("PowerMill.Application");
			} catch (Exception) {
				
				Console.WriteLine("Connection to PM failed!");
				End();
			}
			
			
			string filePath = System.IO.Path.GetDirectoryName( System.Reflection.Assembly.GetExecutingAssembly().Location) +@"\example+excel+tools+and+holder.xlsx";
			
			if (!File.Exists(filePath)) {
				Console.WriteLine(filePath + "     file not found!");
				End();
			}

			
			DataTable excelDataTable = null;
			
			try {
				excelDataTable =LoadWorksheetInDataTable(filePath, GetFirstSheetName(filePath));
			} catch (Exception) {
				
				Console.WriteLine("connection to xllx failed!");
				End();
			}
			
			
			List<ToolDataVAlues> ToolValuesList = new List<ToolDataVAlues>();
			
			foreach (DataRow element in excelDataTable.Rows) {
				List<string> row = new List<string>();
				
				for (int i = 0; i < element.ItemArray.Length; i++) {
					row.Add(element.ItemArray[i].ToString());
				}
				
				ToolValuesList.Add(new ToolDataVAlues(row.ToArray()));
			}
					
			foreach (var element in ToolValuesList) {
				
				using (PMInteraction inter = new PMInteraction(pmApp)) {
					if (inter.Sucess) {
						Console.WriteLine("Connection to PM ok!");
						new Tool(element,pmApp);
						Console.WriteLine("Tool has builded.");
					} else {
						Console.WriteLine("Connection to PM failed!");
					}
				}
				
				
			}
			
			
			End();
		}
		
		static void End() {
			
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
			return;
		}
		
		static string GetFirstSheetName(string fileName)
		{
			DataTable sheetDataAll = new DataTable();
			
			string sheetNamex = null;
		    
		    using (OleDbConnection conn = returnConnection(fileName))
		    {
		       conn.Open();
		       
		       sheetDataAll = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

		        if(sheetDataAll == null)
		        {
		           return null;
		        }
				
		        sheetNamex = sheetDataAll.Rows[0]["TABLE_NAME"].ToString();
		    }
		    
		    return sheetNamex;
		}
		
		static DataTable LoadWorksheetInDataTable(string fileName, string sheetName)
		{           
		    DataTable sheetDataSheet = new DataTable();
		    

		    using (OleDbConnection conn = returnConnection(fileName))
		    {
		       conn.Open();
		       
		       
		        OleDbDataAdapter sheetAdapter = new OleDbDataAdapter("select * from [" + sheetName + "]", conn);
		       
		       sheetAdapter.Fill(sheetDataSheet);
		    }                        
		    return sheetDataSheet;
		}
		
		static OleDbConnection returnConnection(string fileName)
		{
		    return new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;");
		}
		
		
	}
	
	
}