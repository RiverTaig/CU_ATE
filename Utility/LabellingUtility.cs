using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Xsl;

using Miner.Interop;
using Miner.Interop.Process;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ADODB;

namespace Telvent.Designer.Utility
{
    internal static class LabellingUtility
    {
		private const string _ProgID = "SE.Applications.Designer.Utility.LabellingUtility";

		public static string TransformXml(System.Xml.XmlDocument Document, string XslPath)
		{			
			XslCompiledTransform xTrans = null;
			
			#region Do the transform

			try
			{
				bool debugmode = false;
#if DEBUG
				debugmode = true;
#endif
				xTrans = new XslCompiledTransform(debugmode);
				xTrans.Load(XslPath);
			}
			catch (Exception ex)
			{
				throw new Exception("Could not load XSL Transform", ex);
			}

			string results = "";

			try
			{
				System.IO.StringReader SR = new System.IO.StringReader(Document.OuterXml);
				System.Xml.XmlReaderSettings xReadSettings = new System.Xml.XmlReaderSettings();
				xReadSettings.CloseInput = true;
				xReadSettings.ConformanceLevel = System.Xml.ConformanceLevel.Fragment;
				xReadSettings.IgnoreComments = true;
				xReadSettings.IgnoreWhitespace = true;
				System.Xml.XmlReader xReader = System.Xml.XmlReader.Create(SR, xReadSettings);

				StringBuilder SB = new StringBuilder();
				System.IO.StringWriter SW = new System.IO.StringWriter(SB);
				System.Xml.XmlWriterSettings xWriteSettings = new System.Xml.XmlWriterSettings();
				xWriteSettings.CloseOutput = true;
				xWriteSettings.Indent = false;
				xWriteSettings.NewLineOnAttributes = false;
				xWriteSettings.ConformanceLevel = System.Xml.ConformanceLevel.Fragment;

				System.Xml.XmlWriter xWriter = System.Xml.XmlWriter.Create(SW, xWriteSettings);

				xTrans.Transform(xReader, xWriter);

				results = SB.ToString();

				xWriter.Close();
				xReader.Close();

				SR.Dispose();
				SR = null;
				SW.Dispose();
				SW = null;
			}
			catch (Exception ex)
			{
				throw new Exception("Could not transform Design XML.", ex);
			}

			#endregion

			return results;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="PxApp"></param>
		/// <param name="TopList"></param>
		/// <param name="FilterExtent">Optional, if not null it will be used to filter the notes to the current extent</param>
		/// <param name="UseLookupTable">Whether to perform a search / replace of the CUNames in the Design before outputting them</param>
		/// <returns></returns>
        public static string GetConstructionNotes(IMMPxApplication PxApp, IMMPersistentXML ListItem, IEnvelope FilterExtent)
        {
            if (PxApp == null)
                throw new Exception("No Px Application found");
			if (ListItem == null)
				throw new Exception("No item given to generate notes for");

			string XslPath = "";
			try
			{
				XslPath = DesignerUtility.GetPxConfig(PxApp, Constants.PxConfig_ContstructionNotesXslPath);
				if (string.IsNullOrEmpty(XslPath))
					throw new Exception("Obtained an empty reference to the Construction Notes Stylesheet.  Ask your administrator to verify the Px Configuration.");
			}
			catch (Exception ex)
			{
				throw new Exception("Unable to find a Px Configuration for the Construction Notes Stylesheet.  Ask your administrator to verify the Px Configuration.", ex);
			}

			//Our resulting XML must have Work Locations / Gis Units in order to filter it
			System.Xml.XmlDocument modernDocument = null;
			string labelXmlType = DesignerUtility.GetPxConfig(PxApp, Constants.PxConfig_ContstructionNotesXmlSource);
			switch (labelXmlType)
			{
				case "DesignTree":
					modernDocument = GetReportingXml(PxApp, ListItem, LabelXmlType.DesignTree);
					break;
				default:
				case "DesignerXml":
					modernDocument = GetReportingXml(PxApp, ListItem, LabelXmlType.DesignerXml);
					break;
				case "PxXml":
					modernDocument = GetReportingXml(PxApp, ListItem, LabelXmlType.PxXml);
					break;
				case "CostEngine":
					modernDocument = GetReportingXml(PxApp, ListItem, LabelXmlType.CostEngine);
					break;
				case "Custom":
					modernDocument = GetReportingXml(PxApp, ListItem, LabelXmlType.Custom);
					break;
			}

			if (FilterExtent != null)
			{
				#region Fitler the Design Xml

				IRelationalOperator IRO = FilterExtent as IRelationalOperator;

				//Build up a list of Work Locations in the current extent
				List<string> BadWls = new List<string>();
				List<string> BadGus = new List<string>();

				ID8ListItem WlOrCu = null;
				ID8ListItem GivenItem = ListItem as ID8ListItem;
				if (GivenItem == null)
					throw new ApplicationException("Selected item is not a valid list item");

				if (GivenItem.ItemType == mmd8ItemType.mmd8itWorkRequest)
				{
					((ID8List)GivenItem).Reset();
					ID8List Design = ((ID8List)GivenItem).Next(false) as ID8List;
					GivenItem = Design as ID8ListItem;
					((ID8List)GivenItem).Reset();
					WlOrCu = ((ID8List)GivenItem).Next(false);
				}
				else if (GivenItem.ItemType == mmd8ItemType.mmd8itDesign)
				{
					((ID8List)GivenItem).Reset();
					WlOrCu = ((ID8List)GivenItem).Next(false);
				}
				else if (GivenItem.ItemType == mmd8ItemType.mmd8itWorkLocation)
				{
					WlOrCu = (ID8ListItem)GivenItem;
				}
				else
					throw new ApplicationException("Construction notes are not supported on the selected item");

				while (WlOrCu != null)
				{
					if (WlOrCu.ItemType == mmd8ItemType.mmd8itWorkLocation)
					{
						if (!HasD8ChildInExtent(IRO, WlOrCu as ID8List))
							BadWls.Add(((ID8WorkLocation)WlOrCu).ID);
					}
					else
					{
						if (WlOrCu.ItemType == mmd8ItemType.mmitMMGisUnit)
						{
							if (!HasD8ChildInExtent(IRO, WlOrCu as ID8List))
								BadGus.Add(((IMMGisUnit)WlOrCu).GisUnitID.ToString());
						}
					}

					WlOrCu = ((ID8List)GivenItem).Next(false);
				}

				string wlquery = "";
				foreach (string wlid in BadWls)
					if (!string.IsNullOrEmpty(wlid))
						wlquery += "//WORKLOCATION[ID='" + wlid + "']|";
				wlquery = wlquery.TrimEnd("|".ToCharArray());

				string guquery = "";
				foreach (string guid in BadGus)
					if (!string.IsNullOrEmpty(guid))
						guquery += "//GISUNIT[DESIGNER_ID='" + guid + "']|";
				guquery = guquery.TrimEnd("|".ToCharArray());

				string query = wlquery + "|" + guquery;
				query = query.Trim("|".ToCharArray());

				//Filter the xml document to remove the bad wls
				if (!string.IsNullOrEmpty(query))
				{
					foreach (System.Xml.XmlNode BadNode in modernDocument.SelectNodes(query))
						BadNode.ParentNode.RemoveChild(BadNode);
				}

				#endregion
			}

			return TransformXml(modernDocument, XslPath);
        }

		enum LabelXmlType
		{
			DesignTree,
			DesignerXml,
			PxXml,
			CostEngine,
			Custom,
			//DesignerExpress,
		}
		private static System.Xml.XmlDocument GetReportingXml(IMMPxApplication PxApp, IMMPersistentXML ListItem, LabelXmlType XmlType)
		{
			if (PxApp == null)
				throw new Exception("No Px Application found");
			if (ListItem == null)
				throw new Exception("No item given to generate notes for");
			/*
			Stack<ID8ListItem> items = new Stack<ID8ListItem>();
			((ID8List)ListItem).Reset();
			items.Push((ID8ListItem)ListItem);

			int wlCount = 1;
			while (items.Count > 0)
			{
				var item = items.Pop();
				if (item is ID8TopLevel ||
					item is ID8WorkRequest ||
					item is ID8Design)
				{
					((ID8List)item).Reset();
					for (ID8ListItem child = ((ID8List)item).Next(true);
						child != null;
						child = ((ID8List)item).Next(true))
						items.Push(child);
				}
				else if (item is ID8WorkLocation)
				{
					((ID8WorkLocation)item).ID = wlCount.ToString();
					wlCount++;
				}
				else
					continue;
			}
			*/
			switch (XmlType)
			{
				case LabelXmlType.DesignTree:
					{
						Miner.Interop.msxml2.IXMLDOMDocument xDoc = new Miner.Interop.msxml2.DOMDocument();
						ListItem.SaveToDOM(mmXMLFormat.mmXMLFDesign, xDoc);
						
						var newDoc =  new System.Xml.XmlDocument();
						newDoc.LoadXml(xDoc.xml);
						return newDoc;
					}
				default:
				case LabelXmlType.DesignerXml:
					{
						//Hidden packages
						int currentDesign = ((IMMPxApplicationEx)PxApp).CurrentNode.Id;
						return new System.Xml.XmlDocument() { InnerXml = DesignerUtility.GetClassicDesignXml(currentDesign) };
					}
				case LabelXmlType.PxXml:
					{
						//Hidden packages
						int currentDesign = ((IMMPxApplicationEx)PxApp).CurrentNode.Id;
						IMMWMSDesign design = DesignerUtility.GetWmsDesign(PxApp, currentDesign);
						if (design == null)
							throw new Exception("Unable to load design with id of " + currentDesign);

						return new System.Xml.XmlDocument() { InnerXml = DesignerUtility.GetClassicPxXml(design, PxApp) };
					}
				case LabelXmlType.CostEngine:
					{
						//get the px config, load it, run it, return it
						string costEngineProgID = DesignerUtility.GetPxConfig(PxApp, "WMSCostEngine");
						if (string.IsNullOrEmpty(costEngineProgID))
							throw new Exception("Cost Engine Xml Source Defined, but no cost engine is defined");
						Type costEngineType = Type.GetTypeFromProgID(costEngineProgID);
						if (costEngineType == null)
							throw new Exception("Unable to load type for specified cost engine: " + costEngineProgID);

						var rawType = Activator.CreateInstance(costEngineType);
						if (rawType == null)
							throw new Exception("Unable to instantiate cost engine type " + costEngineType);

						IMMWMSCostEngine costEngine = rawType as IMMWMSCostEngine;
						if (costEngine == null)
							throw new Exception("Configured cost engine " + costEngineProgID + " is not of type IMMWMSCostEngine");

						if (!costEngine.Initialize(PxApp))
							throw new Exception("Failed to initialize cost engine");

						return new System.Xml.XmlDocument() { InnerXml = costEngine.Calculate(((IMMPxApplicationEx)PxApp).CurrentNode) };
					}
				case LabelXmlType.Custom:
					throw new Exception("No custom xml reporting source defined");
					/*Or you can reference a custom cost / reporting engine
					CostingEngine.SimpleCostEngine SCE = new Telvent.Applications.Designer.CostingEngine.SimpleCostEngine();
					SCE.Initialize(PxApp);
					CalculatedXml = SCE.Calculate(xDoc.xml);
					*/
					break;
			}

			//Fall through
			return null;
		}

        /// <summary>
        /// Recursively parses a d8list and determines in any of the features
        /// are in the extent.  Intended for use with an ATE.
        /// </summary>
        /// <param name="IRO">Bounding Extent</param>
        /// <param name="List">Designer List Object</param>
        /// <returns></returns>
        private static bool HasD8ChildInExtent(IRelationalOperator IRO, ID8List List)
        {
            bool allchildrenoutofextent = true;

            #region Check the current list item

            if (List is ID8GeoAssoc)
            {
                IFeature GuFeat = ((ID8GeoAssoc)List).AssociatedGeoRow as IFeature;
                if (GuFeat != null && GuFeat.Shape != null)
                {
                    if (!IRO.Disjoint(GuFeat.Shape))
                        allchildrenoutofextent = false;
                }
            }

            #endregion

            List.Reset();
            ID8ListItem Child = List.Next(false);
            while (Child != null && allchildrenoutofextent)
            {
                
                #region Process children until we find a child inside the extent

                if (Child is ID8GeoAssoc)
                {
                    IFeature GuFeat = ((ID8GeoAssoc)Child).AssociatedGeoRow as IFeature;
                    if (GuFeat != null && GuFeat.Shape != null)
                    {
                        if (!IRO.Disjoint(GuFeat.Shape))
                            allchildrenoutofextent = false;
                    }
                }

                if (Child is ID8List)
                    allchildrenoutofextent = !HasD8ChildInExtent(IRO, (ID8List)Child);

                Child = List.Next(false);

                #endregion

            }

            return !allchildrenoutofextent;
        }


		/// <summary>
		/// Executes a query against the give ADODB connection to populate a lookup
		/// dictionary to be used by construction notes.
		/// </summary>
		/// <param name="processConnection">ADODB connection (PX)</param>
		/// <param name="Keys">Column(s) constrained to be unique</param>
		/// <param name="Value">Column to be used as the value</param>
		/// <param name="Table">Table to be queried</param>
		/// <param name="nameLookup">Dictionary to return the results</param>
		private static Dictionary<string, string> QueryLookupTable(ADODB.Connection processConnection, string[] Keys, string Value, string Table)
        {
			Dictionary<string, string> nameLookup = new Dictionary<string, string>();

			if (Keys == null ||
				Keys.Length < 1 ||
				string.IsNullOrEmpty(Value) ||
				string.IsNullOrEmpty(Table))
				return null;

			#region Build the Query

			string columns = string.Join(",", Keys);

            //set up statement
			string sql =
				string.Format("SELECT {0},{1} FROM {2}",
				//ScsConstants.FIELD_WMSCODE,
				//ScsConstants.FIELD_CUNAME,
				//ScsConstants.FIELD_ALTDESC,
				//ScsConstants.TABLE_CU);
			   columns,
			   Value,
			   Table);

			#endregion

            Recordset recordSet = null;
            try
            {

                #region Execute the Query

                try
                {
                    if (processConnection.ConnectionString.Contains(".mdb"))
                        sql = sql.Replace("PROCESS.", "");

                    recordSet = GetRecordSet(processConnection, sql);
                    recordSet.MoveFirst();
                }
                catch (Exception ex)
                {
                    //throw, noting the SQL
                    throw new Exception("Error executing sql statement: " + sql, ex);
                }

                #endregion

                string[] vals = new string[Keys.Length];
                while (!recordSet.EOF)
                {

                    #region Process the results

                    string altdesc = "";
                    try
                    {
                        object oval = null;

                        //The alternate description isn't allowed to be null, this
                        //would cause undesirable results in the construction notes
                        oval = recordSet.Fields[Value].Value;
                        if (oval == null || oval == DBNull.Value)
                            continue;

                        altdesc = oval.ToString();

                        //The pieces in the composite key are allowed to contain 
                        //null values, if they are we give them a special value
                        for (int i = 0; i < Keys.Length; i++)
                        {
                            oval = recordSet.Fields[Keys[i]].Value;
                            if (oval == null || oval == DBNull.Value)
                                vals[i] = "<Null>";
                            else
                                vals[i] = oval.ToString().ToUpper();
                        }
                        /*
                        oval = recordSet.Fields[ScsConstants.FIELD_WMSCODE].Value;
                        if (oval == null || oval == DBNull.Value)
                            wmscode = "<Null>";
                        else
                            wmscode = oval.ToString().ToUpper();

                        oval = recordSet.Fields[ScsConstants.FIELD_CUNAME].Value;
                        if (oval == null || oval == DBNull.Value)
                            name = "<Null>";
                        else
                            name = oval.ToString().ToUpper();
                        */
                        string key = string.Join("\t", vals);
                        if (!nameLookup.ContainsKey(key))
                            nameLookup.Add(key, altdesc);
                    }
                    finally
                    {
                        //move next
                        recordSet.MoveNext();
                    }

                    #endregion

                }
            }
            finally
            {
                if (recordSet != null)
                    recordSet.Close();
                recordSet = null;
            }
            
            return nameLookup;
        }


        private static Recordset GetRecordSet(Connection connection, string sql)
        {
            //create new RecordSet
            ADODB.RecordsetClass rset = new RecordsetClass();

            //open it
            rset.Open(sql, connection, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockOptimistic, 0);

            //return it
            return rset;
        }

    }
}
