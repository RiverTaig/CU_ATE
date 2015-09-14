using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Framework;
using Miner.Interop;
using Miner.Interop.Process;
using ESRI.ArcGIS.Geodatabase;
using Miner.Desktop.DesignerExpress;
using System.Runtime.InteropServices;
using System.Xml;
using ESRI.ArcGIS.esriSystem;

namespace Telvent.Designer.Utility
{
	public static class DesignerUtility
	{
		public static string GetPxConfig(IMMPxApplication PxApp, string ConfigName)
		{
			var config =((IMMPxHelper2)PxApp.Helper).GetConfigValue(ConfigName);
			return config ?? string.Empty;
		}

		public static string GetClassicPxXml(IMMWMSDesign design, IMMPxApplication PxApp)
		{
			object clsDXml = (Activator.CreateInstance(Type.GetTypeFromProgID("mmWorkflowManager.clsDesignerXmlDN")));
			IMMPxXml pxXml = (IMMPxXml)clsDXml;

			//Initialize the above instance with the current application
			pxXml.Initialize(PxApp);

			//Export the xml file for the node with id =  _pxNode.Id and store it in a string
			return pxXml.Export(design.ID);
		}

		public static string GetClassicDesignXml(int designId)
		{
			//Gis MM Package includes full design xml
			//But requires a Designer License to retrieve

			//use normal package
			IMMPackageByUser mpm = new MMPackageManagerClass();
			IMMPackageName pname = new MMPackageNameClass();
			pname.Initialize(-1, designId.ToString(), mmPackageType.mmPTHidden, mmPackageCategory.mmPCDesignXML);
			IMMPackageName packname = mpm.GetPackageNameByUser(pname, designId.ToString());

			if (packname != null)
			{
				IMMPackage package = mpm.GetPackageByUser(packname, designId.ToString());
				if (package != null)
				{

					IMMDesignPackage dpack = package.Contents as IMMDesignPackage;
					return dpack.DesignXML.xml;
				}
			}

			//If we got here, we couldn't load the package
			return null;
		}

		public static string GetDesignXml(IMMWMSDesign design, IMMPxApplication PxApp, bool CompatibilityMode)
		{
			IMMWMSDesign4 dn4 = design as IMMWMSDesign4;
			if (dn4.DesignerProductType == mmWMSDesignerProductType.mmExpress)
				//return GetExpressDesignXml(design, PxApp);
				throw new Exception("Retrieving design Xml for Designer Express designs is not supported.");
			else if (dn4.DesignerProductType == mmWMSDesignerProductType.mmStaker)
				throw new Exception("Retrieving design Xml for Staker designs is not supported.");
			else
			{
				if (CompatibilityMode)
					return GetClassicDesignXml(design.ID);
				else
					return GetClassicPxXml(design, PxApp);
			}
		}

		internal static IApplication GetApplication()
		{
			try
			{
				Type typeOfApp = Type.GetTypeFromProgID("esriFramework.AppRef");
				object objOfApp = Activator.CreateInstance(typeOfApp);
				return (IApplication)objOfApp;
			}
			catch (Exception ex)
			{
				ToolUtility.LogError("Cannot get a reference to a running ArcMap or ArcCatalog instance.", ex);
				
				return null;
			}
		}

		public static IMMPxApplication GetPxApplication()
		{
			ESRI.ArcGIS.Framework.IApplication App = GetApplication();
			if (App == null)
				throw new Exception("Could not load Application Extension");

			IExtension mxExt = App.FindExtensionByName("Workflow Manager Integration");
			IMMPxIntegrationCache pxIntegrationExt = (Miner.Interop.Process.IMMPxIntegrationCache)mxExt;
			IMMPxApplication PxApp = pxIntegrationExt.Application;
			if (PxApp == null)
				throw new Exception("Could not load Px Application Extension");

			return PxApp;
		}

		public static IMMWorkflowManager GetWorkflowManager(IMMPxApplication PxApp)
		{
			return (IMMWorkflowManager)PxApp.FindPxExtensionByName("MMWorkflowManager");
		}

		public static IMMWMSDesign GetWmsDesign(IMMPxApplication PxApp, int DesignId)
		{
			string DesignType = "Design";
			IMMWorkflowManager wfm = GetWorkflowManager(PxApp);
			bool trueVal=true;
			bool falseVal = false;
			return wfm.GetWMSNode(ref DesignType, ref DesignId, ref trueVal, ref trueVal) as IMMWMSDesign;
		}
	}
}
