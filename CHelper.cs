using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Transactions;
using System.Net;
using System.Net.Sockets;
//using ComLib.Logging;
using System.Net.Mail;
using System.Data;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace PDF2CSV
{
	public class CHelper
	{
		public class CNeException : Exception
		{
			public enum ErrorType
			{
				None, Info, Warning, Error, Question
			}

			public ErrorType ErrorTyp;
			public string Msg;

			public CNeException()
			{
				ErrorTyp = CNeException.ErrorType.None;
			}

			public CNeException(Exception exc)
			{
				Msg = exc.Message;
			}

			public CNeException(ErrorType errType, Exception exc)
			{
				Msg = exc.Message;
				ErrorTyp = errType;
			}

			public CNeException(ErrorType errType, string msg)
			{
				ErrorTyp = errType;
				Msg = msg;
			}
		}

		public class CNeErrHandler
		{
			public void HandleErr(CNeException exc)
			{
				string errorType = "";
				MessageBoxIcon msgboxIcon = MessageBoxIcon.None;


				switch (exc.ErrorTyp)
				{
					case CNeException.ErrorType.None:
						break;
					case CNeException.ErrorType.Info:
						errorType = "Information";
						msgboxIcon = MessageBoxIcon.Information;
						break;
					case CNeException.ErrorType.Warning:
						errorType = "Warnung";
						msgboxIcon = MessageBoxIcon.Warning;
						break;
					case CNeException.ErrorType.Error:
						errorType = "Fehler";
						msgboxIcon = MessageBoxIcon.Error;
						break;
					case CNeException.ErrorType.Question:
						errorType = "Frage";
						msgboxIcon = MessageBoxIcon.Question;
						break;
					default:
						break;
				}

				MessageBox.Show(exc.Msg, errorType, MessageBoxButtons.OK, msgboxIcon);
				//MessageBox.Show(string.Format("{0}: {1}",
				//    exc.Msg, string.IsNullOrEmpty(exc.Exc.Message) ? exc.msg.Message : ""),
				//    errorType, MessageBoxButtons.OK, msgboxIcon);
			}

			/*
			public void HandleErr(CNeException exc, LogFile logFile)
			{
				string msg = "";
				if (logFile != null)
				{
					msg += DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + ": ";

					switch (exc.ErrorTyp)
					{
						case CNeException.ErrorType.None:
							break;
						case CNeException.ErrorType.Info:
							msg += "Info - ";
							break;
						case CNeException.ErrorType.Warning:
							msg += "Warning - ";
							break;
						case CNeException.ErrorType.Error:
							msg += "Error - ";
							break;
						case CNeException.ErrorType.Question:
							msg += "Question - ";
							break;
						default:
							break;
					}

					logFile.Message(msg);
				}
			}*/
		}



		public static void CheckSQLConnState(SqlConnection sqlConn)
		{
			if (sqlConn.State != System.Data.ConnectionState.Open)
			{
				sqlConn.Open();
			}
		}

		/// <summary>
		/// Prueft, ob ein gueltiger Zeitraum ausgewaehlt wurde
		/// </summary>
		public static void CheckPeriod(DateTime dtVon, DateTime dtBis)
		{
			if (dtVon == DateTime.MinValue || dtBis == DateTime.MinValue
				|| dtVon > dtBis)
			{
				throw new CHelper.CNeException(CHelper.CNeException.ErrorType.Error, "Bitte einen gültigen Zeitraum auswählen!");
			}
		}


		/// <summary>
		/// Erstellt eine Transaktion ohne Timeout-Zeit
		/// </summary>
		/// <returns></returns>
		public static TransactionScope CreateTransactionScope()
		{
			var transactionOptions = new TransactionOptions();
			transactionOptions.IsolationLevel = System.Transactions.IsolationLevel.ReadCommitted;
			transactionOptions.Timeout = new TimeSpan(0, 5, 0); // 5 Minuten
																													//transactionOptions.Timeout = TransactionManager.MaximumTimeout;
			return new TransactionScope(TransactionScopeOption.Required, transactionOptions);
		}

		public static string SendEMail(string recipient, string bcc, string subject, string bodyMessage)
		{
			string rc = "";

			System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
			SmtpClient client = new SmtpClient("smtp.ionos.de", 587);
			client.EnableSsl = true;
			client.Credentials = new NetworkCredential("VundV@nellinger.com", "V&VB3k0mmtM@il");

			// Set up the email message
			MailMessage message = new MailMessage();
			message.From = new MailAddress("VundV@nellinger.com");
			message.To.Add(recipient);
			if (!string.IsNullOrEmpty(bcc))
			{
				message.Bcc.Add(bcc);
			}
			message.Subject = subject;
			message.Body = bodyMessage;

			try
			{
				// Send the email
				client.Send(message);
			}
			catch (Exception exc)
			{
				rc = exc.Message;
			}

			return rc;
		}

		public static bool IsFileLocked(string filePath)
		{
			FileStream stream = null;
			try
			{
				// Versuche, die Datei exklusiv zu öffnen
				stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
			}
			catch (IOException)
			{
				// Wenn eine IOException auftritt, ist die Datei in Benutzung
				return true;
			}
			finally
			{
				// Schließe den Stream, wenn er geöffnet wurde
				stream?.Close();
			}

			return false;
		}

		public static DateTime GetBuildDT()
		{
			DateTime rc = DateTime.Now;

			try
			{
				rc = System.IO.File.GetLastWriteTime(Application.ExecutablePath);
			}
			catch (Exception)
			{

			}

			return rc;
		}

		public static string SerializeDateTime(DateTime dateTime)
		{
			// Serialisiere DateTime in einen Byte-Array
			using (MemoryStream memoryStream = new MemoryStream())
			{
				BinaryFormatter formatter = new BinaryFormatter();
				formatter.Serialize(memoryStream, dateTime);

				// Konvertiere Byte-Array in Base64-kodierten String
				return Convert.ToBase64String(memoryStream.ToArray());
			}
		}

		public static DateTime DeserializeDateTime(string encodedDateTime)
		{
			// Konvertiere Base64-kodierten String zurück in Byte-Array
			byte[] byteArray = Convert.FromBase64String(encodedDateTime);

			// Deserialisiere Byte-Array zurück in DateTime
			using (MemoryStream memoryStream = new MemoryStream(byteArray))
			{
				BinaryFormatter formatter = new BinaryFormatter();
				return (DateTime)formatter.Deserialize(memoryStream);
			}
		}
	}
}
