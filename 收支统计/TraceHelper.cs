using System;
using System.Diagnostics;

namespace 收支统计
{
	public enum MessageType
	{
		Information,
		Warning,
		Error
	}

	public class TraceHelper
	{
		private static TraceHelper _traceHelper;

		private TraceHelper()
		{
		}

		public static TraceHelper GetInstance()
		{
			if (_traceHelper == null)
			{
				_traceHelper = new TraceHelper();
			}
			return _traceHelper;
		}

		public void Error(string message, string module)
		{
			Log(message, MessageType.Error, module);
		}

		public void Error(Exception ex, string module)
		{
			Log(ex.StackTrace, MessageType.Error, module);
		}

		public void Warning(string message, string module)
		{
			Log(message, MessageType.Warning, module);
		}

		public void Info(string message, string module)
		{
			Log(message, MessageType.Information, module);
		}

		private void Log(string message, MessageType type, string module)
		{
			Trace.WriteLine(string.Format("{0},{1},{2},{3}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), type.ToString(), module, message));
		}
	}
}