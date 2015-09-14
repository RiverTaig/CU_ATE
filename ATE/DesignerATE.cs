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
	/// This base class simplifies the implementation of ATEs that 
    /// inspect the D8TopLevel object.
	/// </summary>
	public abstract class DesignerATE: BaseATE
	{
		protected IApplication App = null;

		public DesignerATE(string Capt, string Msg, string ProgID, string DefaultDisplay)
			: base(Capt, Msg, ProgID, DefaultDisplay)
		{ }	

		/// <summary>
		/// Load the Application, D8TopLevel, and make a call to its own get text.
		/// </summary>
		/// <param name="eTextEvent"></param>
		/// <param name="pMapProdInfo"></param>
		/// <returns></returns>
		protected override string GetText(mmAutoTextEvents eTextEvent, IMMMapProductionInfo pMapProdInfo)
		{
            if (App == null)
				App = DesignerUtility.GetApplication();
			if (App == null)
			{
				ToolUtility.LogError(_ProgID + ": Could not load Application Extension");
				return _defaultDisplay;
			}

            ID8TopLevel topLevel = App.FindExtensionByName("DesignerTopLevel") as ID8TopLevel;
			if (topLevel == null)
				return _defaultDisplay;

			string DisplayString = GetDxText(eTextEvent, pMapProdInfo,topLevel);
			if (string.IsNullOrEmpty(DisplayString))
				return _defaultDisplay;
			else
				return DisplayString;
		}

		/// <summary>
		/// Pass the D8TopLevel to its child
		/// </summary>
		/// <param name="eTextEvent"></param>
		/// <param name="pMapProdInfo"></param>
		/// <param name="topLevel"></param>
		/// <returns></returns>
		protected abstract string GetDxText(mmAutoTextEvents eTextEvent, IMMMapProductionInfo pMapProdInfo, ID8TopLevel topLevel);
	}
}
