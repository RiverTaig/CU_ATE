using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Data;
using System.Xml;

using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.esriSystem;

using log4net;

using Miner.Interop;
using Miner.Interop.Process;
//using Miner.Process;

namespace Telvent.Designer.Utility
{
	public class ToolUtility
	{
		protected static readonly log4net.ILog _logger = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

		public static void LogError(string Message, Exception ex)
		{
			_logger.Error(Message, ex);
		}
		public static void LogError(string Message)
		{
			_logger.Error(Message);
		}

		public static void LogWarning(string Message, Exception ex)
		{
			_logger.Warn(Message, ex);//
		}
		public static void LogWarning(string Message)
		{
			_logger.Warn(Message);
		}
	}
}
