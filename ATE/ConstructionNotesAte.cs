using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using System.Xml.Xsl;

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
	[Guid("81625ADB-A64B-4a9d-8C34-266FA6975B0B")]
	[ClassInterface(ClassInterfaceType.None)]
	[ProgId("SE.Applications.Designer.ConstructionNotesAte")]
	[ComponentCategory(ComCategory.MMCustomTextSources)]
	[ComVisible(true)]
	public class ConstructionNotesAte : DesignerATE
	{
		private string _NoFeatures = "No Features";
		//We could also have an error text, if that was desired

		public ConstructionNotesAte()
			: base("Schneider Electric Construction Notes",
			"Creates Construction Notes for the current page",
			"SE.Applications.Designer.ConstructionNotesAte",
			 "CONSTRUCTIONNOTES")
		{ }

		protected override string GetDxText(mmAutoTextEvents eTextEvent, IMMMapProductionInfo pMapProdInfo, ID8TopLevel topLevel)
		{
			//Only render the ATE when printing/plotting or previewing a print.
			//This boosts performance when working in page layout view
			//without having to 'pause' rendering.
			switch (eTextEvent)
			{
				case mmAutoTextEvents.mmCreate:
				case mmAutoTextEvents.mmDraw:
				case mmAutoTextEvents.mmFinishPlot:
				case mmAutoTextEvents.mmRefresh:
				default:
					return _defaultDisplay;
				case mmAutoTextEvents.mmPlotNewPage:
				case mmAutoTextEvents.mmPrint:
				case mmAutoTextEvents.mmStartPlot:
					break;
			}

			IEnvelope CurrentExtent = null;
			try
			{
				IMap Map = null;
				if (pMapProdInfo != null && pMapProdInfo.Map != null)
					Map = pMapProdInfo.Map;
				else
				{
					//This is requried for print preview or file->export map.
					//During these times there will be no map production object
					//and we will just use the current extent.  Useful for testing
					//and "one-off" maps.
					IApplication App = DesignerUtility.GetApplication();
					IMxDocument MxDoc = App.Document as IMxDocument;
					if (MxDoc == null)
						throw new Exception("Unable to load Map Document from Application");

					Map = MxDoc.FocusMap;
				}

				IActiveView ActiveView = Map as IActiveView;
				if (ActiveView == null)
					throw new Exception("Unable to load Active View from Map");

				CurrentExtent = ActiveView.Extent;
				if (CurrentExtent == null ||
					CurrentExtent.IsEmpty)
					throw new Exception("Unable to determine map extent.");
			}
			catch (Exception ex)
			{
				ToolUtility.LogError("Unable to load extents of the focus map", ex);
				return _defaultDisplay;
			}

			return GetConstructionNotes(topLevel, CurrentExtent);
		}
	
		public string GetConstructionNotes(ID8TopLevel topLevel,IEnvelope FilterExtent)
		{
			//Parse the Design Tree
			ID8List TopList = topLevel as ID8List;
			TopList.Reset();

			IMMPxApplication PxApp = null;			
			try
			{
				PxApp = DesignerUtility.GetPxApplication();
			}
			catch (Exception ex)
			{
				ToolUtility.LogError(_ProgID + ": Could obtain a Px Application reference.", ex);
				return _defaultDisplay;
			}

			try
			{
				TopList.Reset();
				ID8List WorkRequest = TopList.Next(false) as ID8List;
				if (WorkRequest == null)
					return "";

				//WorkRequest.Reset();
				//ID8List Design = WorkRequest.Next(false) as ID8List;

				//Get the Design XML
				IMMPersistentXML WrXml = WorkRequest as IMMPersistentXML;

				string result = Utility.LabellingUtility.GetConstructionNotes(PxApp, WrXml, FilterExtent);
				if (string.IsNullOrEmpty(result))
					return _NoFeatures;
				else
					return result;
			}
			catch (ApplicationException apex)
			{
				//ignore exception, no features
				return _NoFeatures;
			}
			catch (Exception ex)
			{
				ToolUtility.LogError(_ProgID + ": Error Retrieving Construction Notes", ex);
				return _defaultDisplay;
			}
		}

		
	}
}
