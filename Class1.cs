using System;
using System.Threading;

namespace ConnectToAccess
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	class Class1
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
			System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
			conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data source= test.mdb";
			try
			{
				conn.Open();
				Console.WriteLine("Connect To Test.mdb : Success.");
			}
			catch (Exception ex)
			{
				Console.WriteLine("Failed to connect to data source");
			}
			finally
			{
				System.Data.OleDb.OleDbCommand Command1 = new System.Data.OleDb.OleDbCommand("SELECT * FROM D0763422",conn);
				System.Data.OleDb.OleDbDataReader Reader1 = Command1.ExecuteReader();
				while(Reader1.Read())
				{
					Console.WriteLine(Reader1.GetString(1));
					Thread.Sleep(100);
				}
				conn.Close();
				Console.ReadLine();
			}
		}
	}
}
