using System;
using System.IO;
using System.Windows.Forms;

namespace VS.NET_RefeditControl
{
	public class Logger
	{
		private static readonly string CrLf = Environment.NewLine;

		private bool _loggingActive = true;
		private string _logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\RefeditControl\\Logs";
		private string __logFile = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\RefeditControl\\Logs\\logging_.txt";
		private string _logFile
		{
			get
			{
				if (_logFilePath == "" || __logFile == "")
				{
					_logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\RefeditControl\\Logs";
					__logFile = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\RefeditControl\\Logs\\logging_.txt";
				}
				try
				{
					System.IO.Directory.CreateDirectory(_logFilePath);
				}
				catch (Exception ex)
				{
					string msgString = "The RefeditControl log file location (Folder) could not be created at " + CrLf;
					msgString += _logFilePath + CrLf + CrLf + ex.ToString();
					msgString += CrLf + CrLf + "Please go to the RefeditControl options form and nominate a valid location for storing log files.";
					MessageBox.Show(msgString, "Logging Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				return __logFile;

			}
			set
			{
				__logFile = value;
				if (_logFilePath == "" || __logFile == "")
				{
					_logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\RefeditControl\\Logs";
					__logFile = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\RefeditControl\\Logs\\logging_.txt";
					if (_logFilePath.EndsWith("\\"))
					{
						__logFile = _logFilePath + "logging_.txt";
					}
					else
					{
						__logFile = _logFilePath + "\\logging_.txt";
					}
				}
				try
				{
					System.IO.Directory.CreateDirectory(_logFilePath);
				}
				catch (Exception ex)
				{
					string msgString = "The RefeditControl log file location (Folder) could not be created at " + CrLf;
					msgString += _logFilePath + CrLf + CrLf + ex.ToString();
					msgString += CrLf + CrLf + "Please go to the RefeditControl options form and nominate a valid location for storing log files.";
					MessageBox.Show(msgString, "Logging Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}
		public string logFilePath
		{
			get
			{
				if (_logFilePath == "")
				{
					_logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\RefeditControl\\Logs";
				}
				try
				{
					System.IO.Directory.CreateDirectory(_logFilePath);
				}
				catch (Exception ex)
				{
					string msgString = "The RefeditControl log file location (Folder) could not be created at " + CrLf;
					msgString += _logFilePath + CrLf + CrLf + ex.ToString();
					msgString += CrLf + CrLf + "Please go to the RefeditControl options form and nominate a valid location for storing log files.";
					MessageBox.Show(msgString, "Logging Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				return _logFilePath;
			}
			set
			{
				_logFilePath = value;
				if (_logFilePath == "")
				{
					_logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\RefeditControl\\Logs";
				}
				if (_logFilePath.EndsWith("\\"))
				{
					_logFile = _logFilePath + "logging_.txt";
				}
				else
				{
					_logFile = _logFilePath + "\\logging_.txt";
				}
				try
				{
					System.IO.Directory.CreateDirectory(_logFilePath);
				}
				catch (Exception ex)
				{
					string msgString = "The RefeditControl log file location (Folder) could not be created at " + CrLf;
					msgString += _logFilePath + CrLf + CrLf + ex.ToString();
					msgString += CrLf + CrLf + "Please go to the RefeditControl options form and nominate a valid location for storing log files.";
					MessageBox.Show(msgString, "Logging Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		public bool loggingActive
		{
			get
			{
				return _loggingActive;
			}
			set
			{
				_loggingActive = value;
			}
		}

		private static string cleanLog(string logFileName)
		{
			if (logFileName.IndexOf("\\") == -1 && logFileName.IndexOf(".") == -1) return logFileName;
			if (logFileName.IndexOf("\\") != -1)
			{
				logFileName = logFileName.Substring(logFileName.LastIndexOf("\\") + 1);
			}
			if (logFileName.IndexOf(".") != -1)
			{
				logFileName = logFileName.Substring(0, logFileName.IndexOf(".") + 1);
			}
			return logFileName;
		}
		public void LogException(string logFile, string exToString, string errMessage)
		{
			LogException(logFile, exToString, errMessage, "");
		}
		public void LogException(string logFile, string exToString, string errMessage, string ExcelFileName)
		{
			if (!loggingActive) return;
			try
			{
				var w = File.AppendText(_logFile.Replace("logging_.", "logging_" + cleanLog(logFile) + "."));
				w.WriteLine(DateTime.UtcNow.ToShortDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " (" + DateTime.Now.ToLongTimeString() + ") : " + errMessage);
				if (exToString != "") w.WriteLine("Error: " + exToString);
				if (ExcelFileName != "") w.WriteLine("Excel File: " + ExcelFileName);
				w.WriteLine("===============================================================================");
				w.Flush();
				w.Close();
			}
			catch (Exception ex)
			{
				var s = ex.Message;
			}
		}
		public void LogMessage(string logFile, string message, bool breakBefore, bool breakAfter, bool lineBefore, bool lineAfter, bool splitColon)
		{
			StreamWriter w = null;
			try
			{
				w = File.AppendText(_logFile.Replace("logging_.", "logging_" + cleanLog(logFile) + "."));

				if (breakBefore)
				{
					w.WriteLine("");
				}
				if (lineBefore)
				{
					w.WriteLine("===============================================================================");
				}

				if (splitColon)
				{
					if (message.Contains(":"))
					{
						string leftSplit = message.Substring(0, message.IndexOf(":") + 1).Trim();
						string rightSplit = message.Substring(message.IndexOf(":") + 1).Trim();
						leftSplit = leftSplit.PadRight(80);
						message = leftSplit + rightSplit;
					}
				}

				w.WriteLine(DateTime.UtcNow.ToShortDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " (" + DateTime.Now.ToLongTimeString() + ") : " + message);

				if (lineAfter)
				{
					w.WriteLine("===============================================================================");
				}
				if (breakAfter)
				{
					w.WriteLine("");
				}
			}
			catch (Exception ex)
			{
				string s = ex.Message;
			}
			finally
			{
				if (w != null)
				{
					w.Flush();
					w.Close();
				}
			}
		}
	}
}