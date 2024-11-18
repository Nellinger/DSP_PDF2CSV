using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Runtime;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Tabula.Extractors;
using Tabula;
using UglyToad.PdfPig;
using System.Diagnostics.Eventing.Reader;
using System.Reflection;
using System.IO;
using System.Windows.Forms;
using UglyToad.PdfPig.Fonts;
using static OrderPDF2CSV.ConvertPDF2CV4DSP;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using Tabula.Detectors;
using System.Security.Cryptography.X509Certificates;
using System.Security.Policy;
using System.Runtime.Remoting.Contexts;

namespace OrderPDF2CSV
{
	public partial class ConvertPDF2CV4DSP : ServiceBase
	{
		#region UDT

		public enum EnumImportInfosDSP
		{
			None = Int32.MinValue, BestNr = 0, KundenNrEx, Anrede, Name, Vorname, Firma, Adresse1, Adresse2, PLZ, Ort, Land, LandISO, Telefon, Fax, EMail,
			VersandAnrede, VersandName, VersandVorname, VersandFirma, VersandAdresse1, VersandAdresse2, VersandPLZ, VersandOrt,
			VersandLand, VersandLandISO, ArtNr, Preis, Menge, MandantId, Versandart, Lieferscheintyp, UStId, ILN, ExtArticleNo
		};
		public enum EnumParsingMode { NONE = -1, Complete, OnlyIntNumbers, OnlyDoubleNumbers, OnlyLetters }

		public enum EnumBestellStatus
		{
			None = Int32.MinValue, NichtBearbeitet = 0, AusreichenderBestand = 1, KeineArtikelDatenVorhanden = 2,
			KeinBestand = 3, TeilBestand = 4, Gebucht = 5, EnthaeltFehler = 6, TransferOrderInBearbeitung = 7, VorRechnungsDruck = 8, BestellungVersendet = 9,
			BestellungVersandBestaetigungDurchTagesabschluss = 10, BestellungVersandBestaetigungDurchPaketdienstDatenImport = 11
		};

		/// <summary>
		/// Quelle der Bestellung
		/// </summary>
		public enum EnumOrderSource
		{
			NONE = 0, AlbatrosAmazon = 1, WayFair = 2, AmazonDE = 3, AmazonFR = 4, eBayDE = 5, eBayFR = 6, AmazonIT = 7, WebsiteDE = 8,
			WebsiteFRShipDE = 9, GermanCustomerService = 10, EbayDEShipFR = 11, FrenchCS_DE = 12, eBayFRShipDE = 13, AmazonESShipDE = 14, AmazonFRShipDE = 15,
			FrenchCustomerService = 16, WebsiteFR = 17, AmazonES = 18, Plenty = 19, AlbatrosStandard = 20, Manomano_L4F = 21, CDiscountFR_Ship_DE = 22,
			eBayUK, FBA_DE, FBA_ES, FBA_FR, FBA_IT, FBA_UK, AmazonUK, Telesales_UK, Web_UK
		}

		public class CUDTKunde
		{
			public int Id { get; set; }
			public int Nr { get; set; }
			public string NrEx { get; set; }
			public string Anrede { get; set; }
			public string Name { get; set; }
			public string Vorname { get; set; }
			public string Firma { get; set; }
			public string Adresse1 { get; set; }
			public string Adresse2 { get; set; }
			public string AdressePLZ { get; set; }
			public string AdresseOrt { get; set; }
			public string AdresseLand { get; set; }
			public string AdresseLandIso { get; set; }
			public string VersandAnrede { get; set; }
			public string VersandName { get; set; }
			public string VersandVorname { get; set; }
			public string VersandFirma { get; set; }
			public string VersandAdresse1 { get; set; }
			public string VersandAdresse2 { get; set; }
			public string VersandAdressePLZ { get; set; }
			public string VersandAdresseOrt { get; set; }
			public string VersandAdresseLand { get; set; }
			public string VersandAdresseLandIso { get; set; }
			public string Fax { get; set; }
			public string Telefon { get; set; }
			public string EMail { get; set; }
			public string VersandTelefon { get; set; }
			public string VersandFax { get; set; }
			public int Mandant { get; set; }
			public string BestellNr { get; set; }
			public string UStId { get; set; }
			public string ILN { get; set; }

			public CUDTKunde()
			{
				Id = Int32.MinValue;
			}
		}
		public class CUDTBestellung
		{
			public int Id { get; set; }
			public bool NewEntry { get; set; }
			/// <summary>
			/// Soll die benoetigte Menge dieser Bestellung vom Bestand reserviert werden?
			/// </summary>
			public bool UseBestand { get; set; }
			public int PackNr { get; set; }
			public string BestNr { get; set; }
			public int ArtId { get; set; }
			public int OrderNo { get; set; }
			public string ArtNr { get; set; }
			public string ArtDetail { get; set; }
			public int MandantId { get; set; }
			/// <summary>
			/// Menge, die fuer die Bestellung benoetigt wird
			/// </summary>
			public string MengeSoll { get; set; }
			/// <summary>
			/// Menge, die aktuell reserviert ist
			/// </summary>
			public int MengeIst { get; set; }
			/// <summary>
			/// Menge, die aktuell verfuegbar ist
			/// falls MinValue -> kein Bestand vorhanden
			/// </summary>
			public int MengeImBestand { get; set; }
			public double Preis { get; set; }
			public double Gewicht { get; set; }
			public string Versandart { get; set; }
			/// <summary>
			/// Komplette Versandartbez.
			/// </summary>
			public string VersandartStr { get; set; }
			public double Versandkosten { get; set; }
			public string RechnungsNr { get; set; }
			public DateTime DTBestellung { get; set; }
			public DateTime DTTrans { get; set; }
			public DateTime DTVersand { get; set; }
			public double ProvStatus { get; set; } // todo was bedeutet dieses Prop?
			public int PaketArt { get; set; }
			public string ServiceCode { get; set; }
			public string ServiceDescription { get; set; }
			public string ExportData { get; set; }
			public EnumBestellStatus Status { get; set; }
			/// <summary>
			/// Falls ja, wird die benoetigte Menge reserviert (vom Bestand subtrahiert)
			/// Ist true, falls IstMenge > 0
			/// </summary>
			public bool BestellmengeReservieren { get; set; }
			public bool IsWartendeBestellung { get; set; }
			public string LagerOrt { get; set; }
			public string PaketdienstMandantId { get; set; }
			public string ProduktBeschreibung { get; set; }
			/// <summary>
			/// Einzelne Bits stehen fuer versch. Fehler. Siehe Enum ErrorCode
			/// </summary>
			public int ErrorCode { get; set; }
			public string TrackingNos { get; set; }
			public string Bemerkung { get; set; }
			public string AnmerkungKunde { get; set; }
			public bool IsRetoure { get; set; }
			public double Rabatt { get; set; }
			public int Zahlart { get; set; }
			public string ExterneId { get; set; }
			public int Ersatzlieferung { get; set; }
			/// <summary>
			/// Wird nun auch bei Bestellungen gespeichert - fuer Lieferschein
			/// </summary>
			public string ArtBez { get; set; }
			public EnumOrderSource BestellOrt { get; set; }
			public int ExterneItemId { get; set; }
			public int ZeilenNrImportdatei { get; set; }
			/// <summary>
			/// Welcher von mehreren Lieferschein-/Rechungstypen soll bei Bestellungen buchen exportiert werden
			/// </summary>
			public string LieferscheinTyp { get; set; }
			public string ExtArtNr { get; set; }

			public CUDTKunde Kunde { get; set; }

			public CUDTBestellung()
			{
				Id = Int32.MinValue;
				NewEntry = false;
				ArtId = Int32.MinValue;
				MandantId = 28; // DSP
				Status = EnumBestellStatus.None;
				BestellOrt = EnumOrderSource.NONE;

				Kunde = new CUDTKunde();
			}
		}

		public class CUDTSingleFile
		{
			public EnumPDFDocType DocType { get; set; }
			public string SoruceFileNamePDF { get; set; }
			public string SourceFilePathPDF { get; set; }
			public string TargetFilePathPDF { get; set; }
			public CUDTKunde Kunde { get; set; }
			public List<CUDTBestellung> LstBestellungen { get; set; }
			public bool ErrorOccurred { get; set; }

			public CUDTSingleFile()
			{
				LstBestellungen = new List<CUDTBestellung>();
				Kunde = new CUDTKunde();
			}
		}
		#endregion

		private const string FolderInbox = "IN";
		private const string FolderOutbox = "OUT";
		private const string FlagNOK = "ERR";
		private const string FlagOK = "OK";
		private Thread m_ServiceThread;
		private bool m_Stopping;
		private AutoResetEvent m_EvtQuit;
		private AutoResetEvent m_EvtDoTerminate;
		private SqlCommand m_SQLCmd;
		private CSettings m_EntSettings;
		private WaitHandle[] m_EvtArray;
		private ComLib.Logging.LogFile m_Dumpfile;
		private List<string> m_LstLogdata = new List<string>();
		private CHelper.CNeErrHandler m_ErrHandling;
		private const string Header4CSVFile = "BestNr;KdNr;Anrede;Name;Vorname;Firma;Adresse1;Adresse2;PLZ;Ort;Land;Land_ISO;Telefon;Fax;Email;Versand_Anrede;Versand_Name;Versand_Vorname;Versand_Firma;Versand_Adresse1;Versand_Adresse2;Versand_PLZ;Versand_Ort;Versand_Land;Versand_Land_ISO;ArtNr;Preis;Menge;Mandant;Versandart;Lieferschein;Ustid;ILN;ExterneArtNr";
		public List<CUDTSingleFile> m_LstFiles = new List<CUDTSingleFile>();

		private enum EnumTabellenPDFBestellungWebEDI
		{
			NONE = -1, AdresseKaeufer, AdresseLieferant, AdresseLieferung, AdresseRechnung, Lieferdatum,
			Lieferkonditionen, Endempfaenger, BestellteArtikel
		}
		public enum EnumPDFDocType { NONE = 0, OBI_WebEdiWithoutTableEndempfaenger, OBI_WebEdiWithTableEndempfaenger, OBI_Streckenportal, Bauhaus_1 = 100 }

		public ConvertPDF2CV4DSP()
		{
			InitializeComponent();

			SettingsMain settings = new SettingsMain();

			m_EvtDoTerminate = new AutoResetEvent(false);

			//m_Evt_TimeSet = new AutoResetEvent(false);

			m_EvtArray = new WaitHandle[1];
			m_EvtArray[0] = m_EvtDoTerminate;

			// Events an Guis
			m_EvtQuit = new AutoResetEvent(false);
			m_ErrHandling = new CHelper.CNeErrHandler();
		}

		protected override void OnStart(string[] args)
		{
			m_Stopping = false;

			ThreadStart threadStart = new ThreadStart(this.BroadcastThread);

			m_ServiceThread = new Thread(threadStart);
			threadStart = new ThreadStart(this.Run);
			m_ServiceThread = new Thread(threadStart);
			m_Stopping = false;
			m_ServiceThread.Name = "Service main-thread";
			m_ServiceThread.Start();
		}

		protected override void OnStop()
		{
			m_Stopping = true;
			m_EvtDoTerminate.Set();
			m_EvtQuit.Set();

			while ((m_ServiceThread.ThreadState & System.Threading.ThreadState.Stopped) == 0)
			{
				Thread.Sleep(100);
			}

			//m_ErrHandling.Dispose();
		}


		/// <summary>
		/// Diese Methode ist nötig, damit im Debug-Modus OnStart() direkt aus main() gerufen werden kann
		/// </summary>
		/// <param name="args"></param>
		public void OnDebugStart(string[] args)
		{
			OnStart(args);
		}


		/// <summary>
		/// Diese Methode ist nötig, damit im Debug-Modus OnStop() direkt aus main() gerufen werden kann
		/// </summary>
		public void OnDebugStop()
		{
			OnStop();
		}

		/*******************************************************************************
* Hauptprogramm des Dienstes
*******************************************************************************/
		protected void Run()
		{
			int rc;

			try
			{
				SettingsMain settings = new SettingsMain();
				m_Stopping = false;

				int successfullyProcessedFiles = 0;

				if (!CheckLicenseOK())
				{
					m_ErrHandling.HandleErr(new CHelper.CNeException(CHelper.CNeException.ErrorType.Warning, $"Could not find a valid license! " +
									$"{Environment.NewLine}Program is being stopped!"));
					m_Stopping = true;
				}

				{
					while (!m_Stopping)
					{
						try
						{
							string appPath = AppDomain.CurrentDomain.BaseDirectory;
							string[] pdfFiles = Directory.GetFiles(appPath + FolderInbox, "*.pdf");

							if (pdfFiles.Length == 0)
							{
								m_ErrHandling.HandleErr(new CHelper.CNeException(CHelper.CNeException.ErrorType.Warning, $"Could not find any PDF file in folder {appPath + FolderInbox}. " +
									$"{Environment.NewLine}Program is being stopped!"));
								m_Stopping = true;
								continue;
							}

							#region Vorab-Checks, ob geoffnet und anschliessend verschoben werden kann
							List<string> lstLockedFiles = new List<string>();
							foreach (string filePath in pdfFiles)
							{
								if (CHelper.IsFileLocked(filePath))
								{
									lstLockedFiles.Add(filePath);
								}
							}

							if (lstLockedFiles.Count > 0)
							{
								m_ErrHandling.HandleErr(new CHelper.CNeException(CHelper.CNeException.ErrorType.Error,
													$"Please close the following files before processing can start: " +
													$"{Environment.NewLine}'{(string.Join(Environment.NewLine, lstLockedFiles))}'" +
													$"{Environment.NewLine}Program is being stopped!"));
								m_Stopping = true;
								continue;
							}

							#endregion

							List<string> m_FileLines = new List<string>();

							m_LstLogdata.Clear();

							foreach (string filePathSource in pdfFiles)
							{
								CUDTSingleFile currFile = new CUDTSingleFile();
								currFile.SourceFilePathPDF = filePathSource;
								currFile.SoruceFileNamePDF = Path.GetFileName(filePathSource);

								m_LstLogdata.Add($"{Environment.NewLine}############################################{Environment.NewLine}");

								try
								{
									using (PdfDocument document = PdfDocument.Open(filePathSource, new ParsingOptions() { ClipPaths = true }))
									{
										ObjectExtractor oe = new ObjectExtractor(document);
										IExtractionAlgorithm ea = new SpreadsheetExtractionAlgorithm();
										List<Table> tables;
										PageArea pageArea;
										Table currTable;
										string bestNr = "???"; // steht in erster Zeile auf 1. Seite
										DateTime bestDatum = DateTime.MinValue;
										currFile.DocType = EnumPDFDocType.NONE;
										int noOfPages = 0;

										for (int page = 1; page <= 1; page++)
										{
											try
											{
												pageArea = oe.Extract(page);
												noOfPages++;

												if (page == 1) // Bestimmung DocType
												{
													try
													{
														pageArea = oe.Extract(page);
														tables = ea.Extract(pageArea);

														if (currFile.DocType == EnumPDFDocType.NONE) // nur 1x pro Dokument festlegen
														{
															switch (tables.Count)
															{
																case 7:
																	currFile.DocType = EnumPDFDocType.OBI_WebEdiWithoutTableEndempfaenger;
																	break;
																case 8:
																	currFile.DocType = EnumPDFDocType.OBI_WebEdiWithTableEndempfaenger;
																	break;
																case 10:
																	currFile.DocType = EnumPDFDocType.Bauhaus_1;
																	break;
																default:
																	currFile.DocType = EnumPDFDocType.NONE;
																	break;
															}
														}
													}
													catch (Exception)
													{
														// Seite nicht vorhanden
														break;
													}
												}
												else
												{
													m_LstLogdata.Add($"This file contains more than 1 page. Processing will be ignored!");
													currFile.ErrorOccurred = true; // 2 Seiten werden nicht verarbeitet
												}
											}
											catch (Exception)
											{
												break;
											}
										}


										m_LstLogdata.Add($"{DateTime.Now}: Start processing file '{filePathSource}' with a total of {noOfPages} pages...");
										for (int page = 1; page <= 1; page++)
										{
											try
											{
												m_LstLogdata.Add($"Evaluation page #{page}.");

												try
												{
													pageArea = oe.Extract(page);
													tables = ea.Extract(pageArea);
												}
												catch (Exception)
												{
													break;
												}

												m_LstLogdata.Add($"Found {tables.Count} tables on page #{page}.");

												if (currFile.DocType == EnumPDFDocType.OBI_WebEdiWithoutTableEndempfaenger || currFile.DocType == EnumPDFDocType.OBI_WebEdiWithTableEndempfaenger)
												{
													if (page == 1)
													{
														List<string> lstWords = ExtractTextsFromPage(document.GetPage(page));
														if (lstWords.Count >= 4)
														{
															if (lstWords[0].Contains("Bestellung") || lstWords[0].Contains("Order"))
															{
																bestNr = lstWords[1];
																m_LstLogdata.Add($"Found Bestellnummer '{bestNr}'");

																// aktuell ist BestDatum nicht gefordert
																/*
																if (lstWords[2].Contains("vom"))
																{
																	DateTime.TryParse(lstWords[3], out bestDatum);
																}
																*/
															}
															else
															{
																m_LstLogdata.Add($"Could not find field 'Bestellnummer' or 'Order'");
																currFile.ErrorOccurred = true;
															}
														}
														else
														{
															m_LstLogdata.Add($"Could not find value for field 'Bestellnummer' or 'Order'");
															currFile.ErrorOccurred = true;
														}
													}

													for (int tableNo = 0; tableNo < tables.Count; tableNo++)
													{
														if (tableNo <= (int)EnumTabellenPDFBestellungWebEDI.BestellteArtikel)
														{
															m_LstLogdata.Add($"Processing table {tableNo + 1} ({(EnumTabellenPDFBestellungWebEDI)tableNo})...");
															currTable = tables[tableNo];

															if (currTable != null)
															{
																var rows = currTable.Rows;
																m_LstLogdata.Add($"");

																if (rows?.Count > 0)
																{
																	switch ((EnumTabellenPDFBestellungWebEDI)tableNo)
																	{
																		case EnumTabellenPDFBestellungWebEDI.NONE:
																			break;
																		case EnumTabellenPDFBestellungWebEDI.AdresseKaeufer:
																			int idx = 0;
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.ILN), rows[idx++][1], EnumParsingMode.Complete, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.Name), rows[idx++][0], EnumParsingMode.Complete, currFile);
																			if (rows.Count > 8) // Zeile Ansprechpartner ist enthalten und anscheinend optional
																			{
																				GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.Vorname), rows[idx++][0], EnumParsingMode.OnlyLetters, currFile); // teilweise : enthalten wg. verrutschter Spalten
																			}
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.Adresse1), rows[idx++][0], EnumParsingMode.Complete, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.AdressePLZ), rows[idx][0], EnumParsingMode.OnlyIntNumbers, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.AdresseOrt), rows[idx++][0], EnumParsingMode.OnlyLetters, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.AdresseLandIso), rows[idx++][0], EnumParsingMode.Complete, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.Telefon), rows[idx++][0], EnumParsingMode.Complete, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.Fax), rows[idx++][0], EnumParsingMode.Complete, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.UStId), rows[idx++][0], EnumParsingMode.Complete, currFile);
																			break;
																		case EnumTabellenPDFBestellungWebEDI.AdresseLieferant:
																			// KdNr = DSP Lief-Nr inkl. Laenderkuerzel
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.NrEx), rows[5][1], EnumParsingMode.Complete, currFile);
																			break;
																		case EnumTabellenPDFBestellungWebEDI.AdresseLieferung:
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.VersandName), rows[1][0], EnumParsingMode.Complete, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.VersandAdresse1), rows[2][0], EnumParsingMode.Complete, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.VersandAdressePLZ), rows[3][0], EnumParsingMode.OnlyIntNumbers, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.VersandAdresseOrt), rows[3][0], EnumParsingMode.OnlyLetters, currFile);
																			GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currFile.Kunde, nameof(CUDTKunde.VersandAdresseLandIso), rows[4][0], EnumParsingMode.Complete, currFile);
																			break;
																		case EnumTabellenPDFBestellungWebEDI.AdresseRechnung:
																			break;
																		case EnumTabellenPDFBestellungWebEDI.Lieferdatum:
																			break;
																		case EnumTabellenPDFBestellungWebEDI.Lieferkonditionen:
																			break;
																		case EnumTabellenPDFBestellungWebEDI.Endempfaenger:
																		case EnumTabellenPDFBestellungWebEDI.BestellteArtikel:

																			// wenn Endempfaenger vorhanden ist, soll dieser uebersprungen werden
																			if (currFile.DocType == EnumPDFDocType.OBI_WebEdiWithTableEndempfaenger && (EnumTabellenPDFBestellungWebEDI)tableNo == EnumTabellenPDFBestellungWebEDI.Endempfaenger)
																			{
																				break;
																			}

																			// Zweiten Durchlauf verhindern
																			if (currFile.DocType == EnumPDFDocType.OBI_WebEdiWithoutTableEndempfaenger && (EnumTabellenPDFBestellungWebEDI)tableNo == EnumTabellenPDFBestellungWebEDI.BestellteArtikel)
																			{
																				break;
																			}

																			CUDTBestellung currBest = new CUDTBestellung();
																			for (int i = 2; i < rows.Count; i++) // Header direkt ueberspringen
																			{
																				if (i % 2 == 0)
																				{
																					currBest = new CUDTBestellung();
																					currBest.BestNr = bestNr;
																					//currBest.DTBestellung = bestDatum; aktuell nicht relevant

																					GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currBest, nameof(CUDTBestellung.ArtNr), rows[i][1], EnumParsingMode.Complete, currFile);
																					GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currBest, nameof(CUDTBestellung.Preis), rows[i][3], EnumParsingMode.OnlyDoubleNumbers, currFile);
																					GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currBest, nameof(CUDTBestellung.ExtArtNr), rows[i + 1][1], EnumParsingMode.Complete, currFile);
																					GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currBest, nameof(CUDTBestellung.MengeSoll), rows[i + 1][2], EnumParsingMode.OnlyIntNumbers, currFile);
																					currFile.LstBestellungen.Add(currBest);
																				}
																				/*
																				else // Zeile mit Rabatt
																				{
																					// der Rabatt gilt nur auf den 1. Artikel. Falls also mehr als 1x bestellt wird, wird der Rabatt nur auf den 1. Artikel angewendet
																					GetWert4Cell((EnumTabellenPDFBestellungWebEDI)tableNo, currBest, nameof(CUDTBestellung.Preis), rows[i][3], EnumParsingMode.OnlyDoubleNumbers, currFile);

																					
																					if (currBest.Rabatt > 0)
																					{
																						if (currBest.MengeSoll == "1")
																						{
																							currBest.Preis = currBest.Preis - currBest.Rabatt;
																						}
																						else
																						{
																							// neue Bestellung ergaenzen mit der nicht rabattierten Menge
																							CUDTBestellung newBestOhneRabatt = new CUDTBestellung();
																							int mengeSoll = 0;
																							if (int.TryParse(newBestOhneRabatt.MengeSoll, out mengeSoll))
																							{
																								currBest.MengeSoll = "1"; // Menge mit dem Rabatt
																								newBestOhneRabatt.MengeSoll = (mengeSoll - 1).ToString(); // Menge ohne Rabatt

																								// todo Bestell- und Kundendetails kopieren
																								currFile.LstBestellungen.Add(newBestOhneRabatt);
																							}
																						}
																					}
																				}*/
																			}
																			break;
																		default:
																			break;
																	}
																}
																else
																{
																	currFile.ErrorOccurred = true;
																}
															}
														}
													}
												}

												if (currFile.DocType == EnumPDFDocType.OBI_Streckenportal)// Streckenportal
												{
													// aktuell nicht unterstützt
													//throw new Exception($"Document type of file '{filePathSource}' not supported!");

													List<string> lstWords = ExtractTextsFromPage(document.GetPage(page));

													using (PdfDocument document2 = PdfDocument.Open(filePathSource, new ParsingOptions() { ClipPaths = true }))
													{
														ObjectExtractor oe2 = new ObjectExtractor(document2);
														PageArea page2 = oe2.Extract(1);

														// detect canditate table zones
														SimpleNurminenDetectionAlgorithm detector = new SimpleNurminenDetectionAlgorithm();
														var regions2 = detector.Detect(page2);

														IExtractionAlgorithm ea2 = new BasicExtractionAlgorithm();
														List<Table> tables2 = ea2.Extract(page2.GetArea(regions2[0].BoundingBox)); // take first candidate area
														var table2 = tables2[0];
														var rows2 = table2.Rows;
													}

													using (PdfDocument document3 = PdfDocument.Open(filePathSource, new ParsingOptions() { ClipPaths = true }))
													{
														ObjectExtractor oe3 = new ObjectExtractor(document);
														PageArea page3 = oe3.Extract(1);

														IExtractionAlgorithm ea3 = new SpreadsheetExtractionAlgorithm();
														List<Table> tables3 = ea3.Extract(page3);
														var table3 = tables3[0];
														var rows3 = table3.Rows;
													}

													using (PdfDocument document4 = PdfDocument.Open(filePathSource))
													{
														foreach (Page page4 in document4.GetPages())
														{
															GetTextCoordinates(page4, "AB-Solar");

															var words4 = page4.GetWords().ToList();
															var potentialTableRows = IdentifyPotentialTableRows(words4);

															PdfRectangle targetArea = new PdfRectangle(160, 380, 225, 400); // Beispielkoordinaten
															List<string> lstWordsInRectangle = ExtractTextAtPosition(page4, targetArea);

															foreach (var row in potentialTableRows)
															{
																foreach (var word in row)
																{
																	Console.Write($"{word.Text} ");
																}
															}
														}
													}
												}

												if (currFile.DocType == EnumPDFDocType.Bauhaus_1)
												{
													using (PdfDocument document2 = PdfDocument.Open(filePathSource, new ParsingOptions() { ClipPaths = true }))
													{
														ObjectExtractor oe2 = new ObjectExtractor(document2);
														PageArea page2 = oe2.Extract(1);

														// detect canditate table zones
														SimpleNurminenDetectionAlgorithm detector = new SimpleNurminenDetectionAlgorithm();
														var regions2 = detector.Detect(page2);

														// kein schlechter Ansatz: hier wird zeilenweise durch das Dokument iteriert und 4 Spalten des Dokuments sind im Array enthalten
														IExtractionAlgorithm ea2 = new BasicExtractionAlgorithm();
														List<Table> tables2 = ea2.Extract(page2.GetArea(regions2[0].BoundingBox)); // take first candidate area
														var table2 = tables2[0];
														var rows2 = table2.Rows;
													}

													using (PdfDocument document3 = PdfDocument.Open(filePathSource, new ParsingOptions() { ClipPaths = true }))
													{
														ObjectExtractor oe3 = new ObjectExtractor(document);
														PageArea page3 = oe3.Extract(1);

														IExtractionAlgorithm ea3 = new SpreadsheetExtractionAlgorithm();
														List<Table> tables3 = ea3.Extract(page3);
														var table3 = tables3[0];
														var rows3 = table3.Rows;
													}

													using (PdfDocument document4 = PdfDocument.Open(filePathSource))
													{
														foreach (Page page4 in document4.GetPages())
														{
															GetTextCoordinates(page4, "AB-Solar");

															var words4 = page4.GetWords().ToList();
															var potentialTableRows = IdentifyPotentialTableRows(words4);

															PdfRectangle targetArea = new PdfRectangle(160, 380, 225, 400); // Beispielkoordinaten
															List<string> lstWordsInRectangle = ExtractTextAtPosition(page4, targetArea);

															foreach (var row in potentialTableRows)
															{
																foreach (var word in row)
																{
																	Console.Write($"{word.Text} ");
																}
															}
														}
													}
												}
											}
											catch (Exception exc)
											{
												m_Dumpfile.Error($"Error while parsing PDFDocType {currFile.DocType} - file '{currFile.SourceFilePathPDF}': " + exc.Message);
												//m_ErrHandling.HandleErr(new CHelper.CNeException(CHelper.CNeException.ErrorType.Error,
												//$"Error while parsing PDFDocType {currFile.DocType} - file '{currFile.SourceFilePathPDF}': " + exc.Message));
										currFile.ErrorOccurred = true;
											}

											//m_LstLogdata.Add($"Finished processing file and found {lstBestellungen.Count} orders: ");
											// todo Bestellungen und Kundendaten ausgeben
										}
									}
								}
								catch (Exception exc)
								{
									m_ErrHandling.HandleErr(new CHelper.CNeException(CHelper.CNeException.ErrorType.Error,
										$"Error while evaluating file '{filePathSource}': " + exc.Message));
								}

								if (currFile.ErrorOccurred)
								{
									m_LstLogdata.Add($"Error occurred while processing file '{currFile.SourceFilePathPDF}'. Its data will not be part of the output CSV file!");
								}
								else
								{
									successfullyProcessedFiles++;
								}

								m_LstFiles.Add(currFile); // wird hinzugefuegt, auch wenn Fehler aufgetreten sind, um sie spaeter verschieben zu koennen
							}

							// Gesamt-Export
							try
							{
								string targetFolderPath = "";
								string dtStr = DateTime.Now.ToString("yyyyMMdd_HHmmss");
								m_FileLines.Clear();

								m_FileLines.Add(Header4CSVFile);
								foreach (CUDTSingleFile file in m_LstFiles.Where(x => !x.ErrorOccurred).ToList())
								{
									m_FileLines.AddRange(CreateExportData2BestellungsCSV(file.Kunde, file.LstBestellungen));
								}

								foreach (CUDTSingleFile file in m_LstFiles)
								{
									#region Dokument nach Verarbeitung verschieben und Logfile ablegen

									try
									{
										targetFolderPath = Path.GetDirectoryName(file.SourceFilePathPDF);
										targetFolderPath = targetFolderPath.Replace(FolderInbox, FolderOutbox);
										string fileName4PDF = $"{Path.GetFileNameWithoutExtension(file.SourceFilePathPDF)}_{dtStr}_{(file.ErrorOccurred ? FlagNOK : FlagOK)}";

										File.Move(file.SourceFilePathPDF, Path.Combine(targetFolderPath, fileName4PDF + ".pdf"));
									}
									catch (Exception exc)
									{
										m_ErrHandling.HandleErr(new CHelper.CNeException(exc));
									}
									#endregion
								}

								string fileNameBestellungen = $"Bestellungen_{dtStr}.csv";
								File.WriteAllLines(Path.Combine(targetFolderPath, fileNameBestellungen), m_FileLines, Encoding.UTF8);
								string fileNameLogFile = $"Logfile_{dtStr}.txt";
								// Dumpfile schreiben
								m_Dumpfile = new ComLib.Logging.LogFile(fileNameLogFile, Path.Combine(targetFolderPath, fileNameLogFile), DateTime.Now, "", true, 50);
								m_Dumpfile.Settings.AppName = this.ServiceName;
								m_Dumpfile.Log(ComLib.Logging.LogLevel.Debug, "Info");
							}
							catch (Exception exc)
							{
								m_ErrHandling.HandleErr(new CHelper.CNeException(CHelper.CNeException.ErrorType.Error,
								$"Fehler beim Exportieren der CSV-Datei! Das Programm wird beendet... Details = {exc.Message}"));
								m_Stopping = true;
							}
							finally
							{
								m_Dumpfile.Log(ComLib.Logging.LogLevel.Info, "################################# Log messages: ");
								m_Dumpfile.Log(ComLib.Logging.LogLevel.Info, string.Join(Environment.NewLine, m_LstLogdata));
								m_Dumpfile.Flush();
								m_Dumpfile.ShutDown();
								m_Dumpfile = null;
							}

							m_Stopping = true; // Programm beenden
							m_ErrHandling.HandleErr(new CHelper.CNeException(CHelper.CNeException.ErrorType.Info,
								$"{successfullyProcessedFiles} von {pdfFiles.Length} Dokumenten wurden erfolgreich verarbeitet!{Environment.NewLine}Das Programm wird beendet..."));

							rc = WaitHandle.WaitAny(m_EvtArray, 1000, false);

							if (rc != WaitHandle.WaitTimeout)
							{
								switch (rc)
								{
									case 0: // Stop Gateway
													//m_Dumpfile.Write2File("Gateway will be stopped");
										break;

									case 1:
										break;

								}
							}

							// *** zyklische Bearbeitung ***


							// Uhrzeit-Synchronisierung und Lifebeat-Überwachung
							if (!m_Stopping)
							{
								DateTime dtNow = DateTime.Now;

								// Lifebeat
								if (dtNow.Minute == 59 && dtNow.Second < 3)
								{
									m_Dumpfile.Info("Lifebeat...");
								}
							}

#if (DEBUG)
							Console.WriteLine(DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + " - running");
#endif
						}
						catch (Exception e)
						{
							try
							{

								m_Dumpfile.Error(string.Format("Error while processing data! The service will continue! Error: '{0}'!", e.Message));
							}
							catch (Exception)
							{

							}
						}
					} //while
				}
			}
			catch (Exception e)
			{
				m_Dumpfile.Info(string.Format("Fehler in outer tryCatch: '{0}'!", e.Message));
			}
			finally
			{
				//m_Dumpfile.Info(string.Format("Das Programm '{0}' wurde beendet1!", this.ServiceName));
			}

			//m_Dumpfile.Info(string.Format("Das Programm '{0}' wurde beendet2!", this.ServiceName));
		}

		private void GetWert4Cell(EnumTabellenPDFBestellungWebEDI tableNo, object targetObj, string nameOfTargetProperty, Cell cellValue, EnumParsingMode parsingMode, CUDTSingleFile currFile)
		{
			try
			{
				string logData = $"Getting value for table '{tableNo}', Target='{nameOfTargetProperty}', CellData='{cellValue.GetText(false)}'";

				object retrievedValue = new object();
				string numericString = "";
				decimal decimalValue;

				switch (parsingMode)
				{
					case EnumParsingMode.NONE:
						break;
					case EnumParsingMode.Complete:
						retrievedValue = cellValue.GetText(false);
						break;
					case EnumParsingMode.OnlyDoubleNumbers:
						numericString = new string(cellValue.GetText(false).Where(c => char.IsDigit(c) || c == ',').ToArray());
						if (decimal.TryParse(numericString, out decimalValue))
						{
							retrievedValue = decimalValue;
							logData += $" => numeric value {decimalValue}";
						}
						else
						{
							currFile.ErrorOccurred = true;
							logData += "Error: Could not extract decimal value!";
						}
						break;
					case EnumParsingMode.OnlyIntNumbers:
						numericString = new string(cellValue.GetText(false).Where(c => char.IsDigit(c) || c == ',').ToArray());
						if (decimal.TryParse(numericString, out decimalValue))
						{
							try
							{
								int intValue = Convert.ToInt32(decimalValue);
								retrievedValue = intValue;
								logData += $" => numeric value {intValue}";
							}
							catch (Exception)
							{
								currFile.ErrorOccurred = true;
								logData += "Error: Could not convert numeric value to int value!";
							}
						}
						else
						{
							currFile.ErrorOccurred = true;
							logData += "Error: Could not extract numeric value!";
						}
						break;
					case EnumParsingMode.OnlyLetters:
						string onlyLetters = new string(cellValue.GetText(false).Where(c => char.IsLetter(c) || char.IsWhiteSpace(c) || c == '-').ToArray());
						retrievedValue = onlyLetters.Trim();
						break;
					default:
						break;
				}

				SetPropertyValue(targetObj, nameOfTargetProperty, retrievedValue);

				m_LstLogdata.Add(logData);
			}
			catch (Exception exc)
			{
				m_LstLogdata.Add($"Error occurred while parsing: File={currFile.SourceFilePathPDF}, TargetProperty={nameOfTargetProperty} - Value={cellValue}: {exc.Message}");
				currFile.ErrorOccurred = true;
				// dictionary Datei, Ergebnis, Pfade
			}
		}

		private List<string> CreateExportData2BestellungsCSV(CUDTKunde kunde, List<CUDTBestellung> lstBestellungen)
		{
			List<string> fileLines = new List<string>();

			foreach (CUDTBestellung currBest in lstBestellungen)
			{
				#region Datenueberpruefung und -korrektur (Check)

				// Standardwerte
				currBest.VersandartStr = ""; // manuell befuellen
				currBest.LieferscheinTyp = "D"; // 4

				kunde.Name = kunde.Name.Contains("OBIHUB") ? kunde.Name.Replace("OBIHUB", "OBI HUB") : kunde.Name;
				kunde.VersandName = kunde.VersandName.Contains("OBIHUB") ? kunde.VersandName.Replace("OBIHUB", "OBI HUB") : kunde.VersandName;
				kunde.AdresseLandIso = kunde.AdresseLandIso == "A" ? "AT" : kunde.AdresseLandIso;
				kunde.VersandAdresseLandIso = kunde.VersandAdresseLandIso == "A" ? "AT" : kunde.VersandAdresseLandIso;
				kunde.AdresseLandIso = kunde.AdresseLandIso == "D" ? "DE" : kunde.AdresseLandIso;
				kunde.VersandAdresseLandIso = kunde.VersandAdresseLandIso == "D" ? "DE" : kunde.VersandAdresseLandIso;
				kunde.AdresseLandIso = kunde.AdresseLandIso == "SL" ? "SVK" : kunde.AdresseLandIso;
				kunde.VersandAdresseLandIso = kunde.VersandAdresseLandIso == "SL" ? "SVK" : kunde.VersandAdresseLandIso;

				switch (kunde.AdresseLandIso)
				{
					case "DE":
						kunde.AdresseLand = "Germany";
						break;
					case "AT":
						kunde.AdresseLand = "Austria";
						break;
					case "SVK":
						kunde.AdresseLand = "Slovakia";
						break;
					default:
						break;
				}

				switch (kunde.VersandAdresseLandIso)
				{
					case "DE":
						kunde.VersandAdresseLand = "Germany";
						break;
					case "AT":
						kunde.VersandAdresseLand = "Austria";
						break;
					case "SVK":
						kunde.VersandAdresseLand = "Slovakia";
						break;
					default:
						break;
				}

				#endregion

				List<string> currLine = new List<string>();
				currLine.Add(currBest.BestNr);
				currLine.Add(kunde.NrEx);
				currLine.Add(kunde.Anrede);
				currLine.Add(kunde.Name);
				currLine.Add(kunde.Name);
				currLine.Add(kunde.Name);
				currLine.Add(kunde.Adresse1);
				currLine.Add(kunde.Adresse2);
				currLine.Add(kunde.AdressePLZ);
				currLine.Add(kunde.AdresseOrt);
				currLine.Add(kunde.AdresseLand);
				currLine.Add(kunde.AdresseLandIso);
				currLine.Add(kunde.Telefon);
				currLine.Add(kunde.Fax);
				currLine.Add(kunde.EMail);

				currLine.Add(kunde.VersandAnrede);
				currLine.Add(kunde.VersandName);
				currLine.Add(kunde.VersandName);
				currLine.Add(kunde.VersandName);
				currLine.Add(kunde.VersandAdresse1);
				currLine.Add(kunde.VersandAdresse2);
				currLine.Add(kunde.VersandAdressePLZ);
				currLine.Add(kunde.VersandAdresseOrt);
				currLine.Add(kunde.VersandAdresseLand);
				currLine.Add(kunde.VersandAdresseLandIso);

				currLine.Add(currBest.ArtNr);
				currLine.Add(currBest.Preis.ToString("F2"));
				currLine.Add(currBest.MengeSoll);
				currLine.Add(currBest.MandantId.ToString());
				currLine.Add(currBest.VersandartStr);
				currLine.Add(currBest.LieferscheinTyp);
				currLine.Add(kunde.UStId);
				currLine.Add(kunde.ILN);
				currLine.Add(currBest.ExtArtNr);

				fileLines.Add(string.Join(";", currLine));
			}

			return fileLines;
		}

		public static void SetPropertyValue(object obj, string propertyName, object value)
		{
			PropertyInfo prop = obj.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
			if (prop != null && prop.CanWrite)
			{
				if (value == null)
				{
					if (Nullable.GetUnderlyingType(prop.PropertyType) != null || !prop.PropertyType.IsValueType)
					{
						prop.SetValue(obj, null, null);
					}
					else
					{
						throw new ArgumentException($"Cannot set null to non-nullable property '{propertyName}'.");
					}
				}
				else
				{
					prop.SetValue(obj, Convert.ChangeType(value, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType), null);
				}
			}
			else
			{
				throw new ArgumentException($"Property '{propertyName}' not found or not writable.");
			}
		}

		static void GetTextCoordinates(Page page, string searchWord)
		{
			List<string> lstWords = new List<string>();
			var words = page.GetWords();
			foreach (var word in words)
			{
				lstWords.Add(word.Text);
				if (word.Text.ToUpper().Contains(searchWord.ToUpper()))
				{
					string koordinaten = $"Position: Left={word.BoundingBox.Left}, Bottom={word.BoundingBox.Bottom}, Right={word.BoundingBox.Right}, Top={word.BoundingBox.Top}";
					break;
				}
			}

			lstWords.Sort();
		}

		private List<string> ExtractTextsFromPage(Page page)
		{
			List<string> rc = new List<string>();

			var anno = page.GetMarkedContents();

			var textRegions = new List<TextRegion>();
			var words = page.GetWords();

			foreach (var word in words)
			{
				rc.Add(word.Text);
			}

			return rc;
		}

		public class TextRegion
		{
			public string Text { get; set; }
			//public PdfRectangle BoundingBox { get; set; }
		}

		private List<List<Word>> IdentifyPotentialTableRows(List<Word> words)
		{
			var rows = new List<List<Word>>();
			words = words.OrderBy(w => w.BoundingBox.Bottom).ThenBy(w => w.BoundingBox.Left).ToList();

			List<Word> currentRow = new List<Word>();
			double currentBottom = words[0].BoundingBox.Bottom;

			foreach (var word in words)
			{
				if (Math.Abs(word.BoundingBox.Bottom - currentBottom) < 5) // Adjust threshold as needed
				{
					currentRow.Add(word);
				}
				else
				{
					rows.Add(currentRow);
					currentRow = new List<Word> { word };
					currentBottom = word.BoundingBox.Bottom;
				}
			}
			rows.Add(currentRow);

			return rows;
		}

		static List<string> ExtractTextAtPosition(Page page, PdfRectangle targetArea)
		{
			List<string> rc = new List<string>();
			var words = page.GetWords();
			foreach (var word in words)
			{
				if (IsWordInRectangle(word, targetArea))
				{
					rc.Add(word.Text);
				}
			}

			return rc;
		}

		static bool IsWordInRectangle(Word word, PdfRectangle rectangle)
		{
			return word.BoundingBox.Left >= rectangle.Left &&
						 word.BoundingBox.Right <= rectangle.Right &&
						 word.BoundingBox.Bottom >= rectangle.Bottom &&
						 word.BoundingBox.Top <= rectangle.Top;
		}

		/// <summary>
		/// Prueft die Lizenz, falls die Lizenzdatei vorhanden ist.
		/// Wurde sie geloescht, ist die Lizenz gueltig
		/// </summary>
		/// <returns></returns>
		private bool CheckLicenseOK()
		{
			// todo auf Lizenz-Datei pruefen, dann ist Rest egal
			//return true;
			bool rc = true;

			try
			{
				string path4Check = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\ParsingInfo.obj";
				if (!File.Exists(path4Check))
				{
					rc = true;
				}
				else
				{
					if ((DateTime.Now - CHelper.GetBuildDT()).TotalDays > 45)
					{
						rc = false;
					}
				}
			}
			catch (Exception exc)
			{
				m_ErrHandling.HandleErr(new CHelper.CNeException(CHelper.CNeException.ErrorType.Error, "Error while checking license: " + exc.Message));
				rc = false;
			}

			return rc;
		}


		protected void BroadcastThread()
		{

		}
	}
}
