/*
 * Created by SharpDevelop.
 * User: Ondra
 * Date: 25/03/2016
 * Time: 16:15
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;



namespace SpainPMExcelImport
{
	/// <summary>
	/// Description of PMInteraction.
	/// </summary>
	public class PMInteraction : IDisposable
	{
		bool sucess = true;
		public bool Sucess {
			get {
				return sucess;
			}
		}
		

		bool _dialogy;
		bool _disableAll = true;
		
		bool _pmConnectionButton = false;
		
		PowerMILL.Application _app = null;


		public PMInteraction(PowerMILL.Application oApplication) 
		{
			_pmConnectionButton = false;
			_app = oApplication;
			
			if (_app==null) {
				sucess=false;
				return;
			}
			
			Console.WriteLine("SpinWait 5s timeout.........");
			System.Threading.SpinWait.SpinUntil(()=>!_app.Busy,5000);
			Console.WriteLine("SpinWait end");
			if (_app.Busy) {
				sucess=false;
				return;
			}

			_app.DoCommand(@"QUIT");
			_app.DoCommand(@"QUIT");
			_app.DoCommand(@"QUIT");
			_app.DoCommand(@"ECHO OFF DCPDEBUG UNTRACE COMMAND ACCEPT");
			
			_dialogy = getActualDialogs(_app);
			_app.DoCommand("DIALOGS MESSAGE OFF");
			_app.DoCommand("DIALOGS ERROR OFF");
			
			
			
		}
		
		public bool getActualDialogs(PowerMILL.Application oAppliacation)
		{
 			int err = 0;
 			string vysledek = null;
 			oAppliacation.ExecuteEx(@"print $Status.Dialog.Message",out err,out vysledek);
 			vysledek = vysledek.Replace(Environment.NewLine,"").Trim();
			return vysledek == "1";
 			
		}
		

		
		
		#region IDisposable implementation
		public void Dispose()
		{

			
			if (_dialogy) {
				if (_pmConnectionButton) {
					_app.DoCommand("DIALOGS MESSAGE ON");
					_app.DoCommand("DIALOGS ERROR ON");
				} else {
					_app.DoCommand("DIALOGS MESSAGE ON");
					_app.DoCommand("DIALOGS ERROR ON");
				}
			}
			


		}
		#endregion
	}
}
