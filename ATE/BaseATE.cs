using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;

using Miner.Interop;
using Miner.Interop.Process;
using Miner.ComCategories;

using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.esriSystem;
using log4net;
using Telvent.Designer.Utility;

namespace Telvent.Designer.ATE
{
	/// <summary>
	/// Base class for all custom ATEs in this solution
	/// </summary>
	public abstract class BaseATE : IMMAutoTextSource
	{
		//Set the internal variables
		public BaseATE(string Capt, string Msg, string ProgID, string DefaultDisplay)
		{
			_Caption = Capt;
			_Message = Msg;
			_ProgID = ProgID;
			_defaultDisplay = DefaultDisplay;
		}


		#region Component Category Registration
		// Function Executed during the Registration process
		[ComRegisterFunction()]
		static void RegisterFunction(String sKey)
		{
			MMCustomTextSources.Register(sKey);
		}

		//Function executed during the UnRegistration process
		[ComUnregisterFunction()]
		static void UnregisterFunction(String sKey)
		{
			MMCustomTextSources.Unregister(sKey);
		}
		#endregion#

		#region IMMAutoTextSource Members

		string IMMAutoTextSource.Caption
		{
			get { return _Caption; }
		}

		string IMMAutoTextSource.Message
		{
			get { return _Message; }
		}

		bool IMMAutoTextSource.NeedRefresh(mmAutoTextEvents eTextEvent)
		{
			return true;
		}

		string IMMAutoTextSource.ProgID
		{
			get { return _ProgID; }
		}

		string IMMAutoTextSource.TextString(mmAutoTextEvents eTextEvent, IMMMapProductionInfo pMapProdInfo)
		{
			string sReturnVal = _defaultDisplay;

			try
			{
				sReturnVal = GetText(eTextEvent, pMapProdInfo);
			}
			catch (Exception ex)
			{
				ToolUtility.LogError("Error getting autotext", ex);
			}

			if (string.IsNullOrEmpty(sReturnVal))
				return _NullString;
			else
				return sReturnVal;
		}

		#endregion

		#region Abstract members
		protected string _defaultDisplay;
		protected string _Caption;
		protected string _Message;
		protected string _ProgID;
		//return a space if no data is found. This ensures the ATE is still printed as a null or empty string will cause it to not be printed
		protected string _NullString = " ";

		protected abstract string GetText(mmAutoTextEvents eTextEvent, IMMMapProductionInfo pMapProdInfo);
		#endregion

	}
}
