using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;

namespace PDF2CSV
{
	public class CSettings : IDisposable
	{
		/// <summary>
		/// Vergebener Int-Wert entspricht der DB-Id
		/// </summary>
		public enum EnumSettingsParam
		{
			PathSuffix4ExportDPD = 2, PathSuffix4ExportDHL, PathSuffix4ExportGLS, PathSuffix4Warensendung,
			PathRoot4Export, FolderNameExportArchive, FolderName4Invoice, FolderName4VersanddatenExport,
			Path2ReportFolder, Path2ArtikelImportLogfile, Path2PDFMergePrg, FolderName4VersanddatenExport_Lights4Fun,
			FolderName4VersanddatenExport_WeineDe, FolderName4VersanddatenExport_Albatross,
			FolderName4BestandExport_Lights4Fun, FolderName4BestandExport_WeineDe, FolderName4BestandExport_Albatross,
			Folder4AutomaticTagesabschlussImport, Folder4VersanddatenExport_BlackRock, Folder4BestandExport_BlackRock,
			FolderName4BestandExport_NetProfi, FolderName4VersanddatenExport_NetProfi, CheckEditRights, FolderName4Packliste,
			FolderName4ImportBestellungenLog, Folder4VersanddatenExport_Makstore, PathSuffix4Bulky,
			PathReportExportL4F, PathReportExportAlbatros, PathReportExportWeineDe, PathSuffix4ExportUPS, Folder4VersanddatenExport_Loadhog,
			Folder4BestanddatenExport_Loadhog, Folder4BestanddatenExport_Medneo, Folder4BestanddatenExport_ROCS, Folder4VersanddatenExport_Medneo,
			Folder4VersanddatenExport_ROCS, Folder4BBestanddatenExport_ATC, Folder4VersanddatenExport_ATC, Folder4ItemFulfilmentExport_L4F,
			Folder4InventoryAdjustmentExport_L4F, Folder4ItemFulfilmentExport_Sopost, Folder4InventoryAdjustmentExport_Sopost,
			FolderName4BestandExport_TourMade, FolderName4VersanddatenExport_TourMade, FolderName4WartendeExport_WeineDe,
			FolderName4DocExport, Folder4BestanddatenExport_CSTrading, Folder4BestanddatenExport_NRTH50,
			Folder4VersanddatenExport_CSTrading, Folder4VersanddatenExport_NRTH50, Folder4AutomaticTagesabschlussWaproImport,
			Folder4BestanddatenExport_NAMWood, Folder4VersanddatenExport_NAMWood,
			Folder4ImportBestellungen_NRTH50, Folder4ImportBestellungen_Albatros, Folder4ImportBestellungen_CSTrading,
			Folder4ImportBestellungen_L4F, Folder4ImportBestellungen_Medneo, Folder4ImportBestellungen_MySportsWorld, Folder4ImportBestellungen_TourMade, Folder4ImportBestellungen_Weine,
			Folder4ImportBestellungen_GTSE = 64, FolderName4BestandExport_GTSE = 65, FolderName4VersanddatenExport_GTSE = 66, Folder4ItemFulfilmentExport_GTSE = 67,
			Folder4ImportBestellungen_CopyCat = 68, FolderName4BestandExport_CopyCat = 69, FolderName4VersanddatenExport_CopyCat = 70,
			FolderName4VersanddatenExport_WeineDe_UPS,
			Folder4ImportBestellungen_DSP = 72, FolderName4BestandExport_DSP = 73, FolderName4VersanddatenExport_DSP = 74
		};

		public enum EnumVersandarten
		{
			NONE = 0, DPD_Standard, PostWarensendung, Spedition, DPD_Nachnahme, Post_Brief, Post_Brief_International, Post_Einschreiben_National,
			Post_Einschreiben_International, DPD_Express_0830, DPD_Express_1200, DPD_Express_1800, DPD_Samstag_Express, DHL_Standard,
			DHL_Nachnahme, GLS_Standard, GLS_Nachnahme, UPS_Standard, DPD_ParcelLetter, DPD_Express_1000, DHL_Bulky_DOM, DHL_DP_LANG,
			Asendia = 30, Amazon_DropShip = 40, Abholung = 50
		}

		public class CUDTSettings
		{
			public int Id;
			public string Beschreibung;
			public string Value;
			public EnumSettingsParam EnumType;
		}


		private Dictionary<EnumSettingsParam, CUDTSettings> m_DictSettings;
		private SqlConnection m_SQLConnection;
		private SqlCommand m_SQLCmd;

		public CSettings(SqlConnection sqlConn)
		{
			m_SQLConnection = sqlConn;
			m_SQLCmd = new SqlCommand();
			m_SQLCmd.Connection = m_SQLConnection;

			m_DictSettings = new Dictionary<EnumSettingsParam, CUDTSettings>();
		}

		public void GetAllData()
		{
			GetAllSettings();
		}


		private void GetAllSettings()
		{
			CHelper.CheckSQLConnState(m_SQLConnection);

			m_SQLCmd.CommandText = "SELECT Id, Bezeichnung, Value FROM dbo.Settings";

			try
			{
				SqlDataReader r = m_SQLCmd.ExecuteReader();
				CUDTSettings setting;
				m_DictSettings.Clear();

				while (r.Read())
				{
					setting = new CUDTSettings();
					int idx = -1;

					setting.Id = r.IsDBNull(++idx) ? Int32.MinValue : r.GetInt32(idx);
					setting.Beschreibung = r.IsDBNull(++idx) ? "" : r.GetString(idx);
					setting.Value = r.IsDBNull(++idx) ? "" : r.GetString(idx);
					// OK, wenn in Entity gecastet wird?
					setting.EnumType = (EnumSettingsParam)setting.Id;

					m_DictSettings.Add(setting.EnumType, setting);
				}
			}
			catch (Exception exc)
			{
				throw new CHelper.CNeException(CHelper.CNeException.ErrorType.Error, exc.Message);
			}
			finally
			{
				m_SQLConnection.Close();
			}
		}

		#region Props

		public Dictionary<EnumSettingsParam, CUDTSettings> DictSettings
		{
			get
			{
				return m_DictSettings;
			}
		}


		#endregion

		public void Dispose()
		{
			m_DictSettings.Clear();
			m_DictSettings = null;
		}
	}
}
