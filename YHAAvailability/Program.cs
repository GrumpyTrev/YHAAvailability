using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Net.Http;
using HtmlAgilityPack;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace YHAAvailability
{
	class Program
	{
		static void Main( string[] args )
		{
			MonthRequest[] monthRequests = { new MonthRequest( "2019-03-01", "Mar" ), new MonthRequest( "2019-04-01", "Apr" ),
				new MonthRequest( "2019-05-01", "May" ), new MonthRequest( "2019-06-01", "Jun" ), new MonthRequest( "2019-07-01", "Jly" ),
				new MonthRequest( "2019-08-01", "Aug" ), new MonthRequest( "2019-09-01", "Sep" ), new MonthRequest( "2019-10-01", "Oct" ) };

			List<HostelGroup> groups = new List<HostelGroup> {
				new HostelGroup( "Peak District" ).AddHostel( new Hostel( "Ilam Hall", 118 ), new Hostel( "Hartington Hall", 104 ),
				new Hostel( "Edale", 81 ), new Hostel( "Castleton", 52 ), new Hostel( "Hathersage", 106 ), new Hostel( "Eyam", 93 ), new Hostel( "Ravenstor", 195 ),
				new Hostel( "Youlgreave", 248 ) ),

				new HostelGroup( "Pembrokeshire" ).AddHostel( new Hostel( "Broad Haven", 35 ), new Hostel( "St Davids", 200 ),
				new Hostel( "Pwll Deri", 193 ), new Hostel( "Newport", 254 ), new Hostel( "Poppit Sands", 190 ), new Hostel( "Manorbier", 167 ) ),

				new HostelGroup( "Lakes" ).AddHostel( new Hostel( "Skiddaw", 745 ), new Hostel( "Keswick", 129 ),
				new Hostel( "Buttermere", 40 ), new Hostel( "Borrowdale", 157 ), new Hostel( "Ennerdale", 88 ), new Hostel( "Black Sail", 21 ),
				new Hostel( "Honister Hause", 115 ), new Hostel( "Helvellyn", 111 ), new Hostel( "Patterdale", 182 ), new Hostel( "Ambleside", 4 ),
				new Hostel( "Grasmere", 98 ), new Hostel( "Langdale", 112 ), new Hostel( "Coniston", 62 ), new Hostel( "Coniston Copper", 61 ),
				new Hostel( "Eskdale", 90 ), new Hostel( "Wasdale Hall", 234 ) ),

				new HostelGroup( "Yorkshire Dales" ).AddHostel( new Hostel( "Osmotherley", 757 ), new Hostel( "Helmsley", 110 ),
				new Hostel( "Whitby", 337 ), new Hostel( "Scarborough", 767 ), new Hostel( "Boggle Hole", 24 ), new Hostel( "Dalby Forest", 148 ),
				new Hostel( "York", 247 ) )
			};

			foreach ( HostelGroup group in groups )
			{
				int hostelCount = 0;

				foreach ( Hostel hostel in group.Hostels )
				{
					foreach ( MonthRequest request in monthRequests )
					{
						// Get the month availability for the hostel
						HtmlDocument doc = new HtmlDocument();
						doc.LoadHtml( MakeRequest( hostel.Id.ToString( "D6" ), request.Request ).Result );

						HtmlNodeCollection dormNodes = doc.DocumentNode.SelectNodes( "//div[@class='dates dorm']" );

						if ( dormNodes.Count > 0 )
						{
							List<string> dates = dormNodes[0].Descendants( "span" ).Select( data => data.Attributes[ "data-date" ].Value ).ToList();
							List<string> availabilities = null;

							if ( dormNodes.Count > 1 )
							{
								availabilities = dormNodes[ 1 ].Descendants( "span" ).Select( data => data.ChildNodes[ "#text" ].InnerText ).ToList();
							}

							// Parse and save to the AvailabilityData 
							SaveAvailabiltyForHostelMonth( dates, availabilities, group.Availability, hostelCount );
						}
					}

					hostelCount++;
				}
			}

			// Create the spreadsheet
			GenerateSpreadsheet( @"C:\temp\YHAResponse.xlsx", groups );

			Console.ReadKey();
		}

		/// <summary>
		/// Request the availability for the specified hostel and month
		/// </summary>
		/// <param name="hostel"></param>
		/// <param name="month"></param>
		/// <returns></returns>
		static async Task<string> MakeRequest( string hostel, string month )
		{
			HttpResponseMessage response = await client.PostAsync( "https://availabilitycalendar.yha.org.uk/availabilityCalendar.php", 
				new FormUrlEncodedContent( new Dictionary<string, string> { { "house", hostel }, { "viewdate", month }, { "filter", "dorm" },
					{ "males", "1" }, { "type", "020" } } ) );

			return await response.Content.ReadAsStringAsync();
		}

		/// <summary>
		/// Parse the response data for a particual month and add it to the availability data
		/// </summary>
		/// <param name="responseData"></param>
		/// <param name="monthName"></param>
		/// <param name="availability"></param>
		static void SaveAvailabiltyForHostelMonth( List<string> dates, List<string> availabilities, AvailabilityData availability, int hostelCount )
		{
			// Enumerate the results
			int index = 0;
			List<string>.Enumerator enumerator = dates.GetEnumerator();
			while ( enumerator.MoveNext() == true )
			{
				string date = enumerator.Current;

				// Format is 'yyyy-mm-dd'
				if ( date.Length == 10 )
				{
					DateTime fullDate = DateTime.Parse( date );

					// Add an entry for this date to the dictionary if not already there
					if ( availabilities != null )
					{
						availability.AddAvailability( 
							string.Format( "{0}, {1} {2}, 2019", fullDate.DayOfWeek.ToString().Substring( 0, 3 ), fullDate.ToString( "MMM" ), fullDate.Day ), 
							( availabilities[ index ].Length > 6 ) && ( availabilities[ index ][ 0 ] == '&' ), hostelCount );
					}
					else
					{
						availability.AddAvailability(
							string.Format( "{0}, {1} {2}, 2019", fullDate.DayOfWeek.ToString().Substring( 0, 3 ), fullDate.ToString( "MMM" ), fullDate.Day ), false, hostelCount );
					}
				}

				index++;
			}
		}

		static void GenerateSpreadsheet( string fileName, List<HostelGroup> groups )
		{
			using ( SpreadsheetDocument document = SpreadsheetDocument.Create( fileName, SpreadsheetDocumentType.Workbook ) )
			{
				// Add a Workbook container to the document and put a workbook in it
				WorkbookPart workbookPart = document.AddWorkbookPart();
				workbookPart.Workbook = new Workbook();

				// Add an initially empty collection of sheets to the workbook
				Sheets sheets = workbookPart.Workbook.AppendChild( new Sheets() );

				// Adding style
				WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
				stylePart.Stylesheet = GenerateStylesheet();
				stylePart.Stylesheet.Save();

				// Create and populate a worksheet per group
				uint sheetId = 1;
				foreach ( HostelGroup group in groups )
				{
					// Add a worksheet container to the workbook container, and put a worksheet in it
					WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
					worksheetPart.Worksheet = new Worksheet();

					// Setting up columns collection and add the date column format.
					// Columns are worksheet specific and cannot be shared between multiple worksheets
					Columns columns = new Columns( new Column { Min = 1, Max = 1, Width = 18, CustomWidth = true } );
					worksheetPart.Worksheet.AppendChild( columns );

					// Create a sheet to contain the data (well at least a link to the data) and append it to the collection held by the workbook
					sheets.Append( new Sheet() { Id = workbookPart.GetIdOfPart( worksheetPart ), SheetId = sheetId++, Name = group.Name } );

					// Save the workbook
					workbookPart.Workbook.Save();

					// Constructing header of hostel names
					Row row = new Row();
					row.Append( ConstructCell( " ", CellValues.String, 1 ) );
					uint columnNo = 2;

					foreach ( Hostel hostel in group.Hostels )
					{
						row.Append( ConstructCell( hostel.Name, CellValues.String, 1 ) );
						columns.Append( new Column { Min = columnNo, Max = columnNo++, Width = hostel.Name.Length + 1, CustomWidth = true } );
					}

					// Now actually create somewhere to store the data
					SheetData sheetData = worksheetPart.Worksheet.AppendChild( new SheetData() );

					// Insert the header row to the Sheet Data
					sheetData.AppendChild( row );

					// Inserting each date
					foreach ( string date in group.Availability.Dates )
					{
						row = new Row();
						row.Append( ConstructCell( date, CellValues.String, 4 ) );

						// And now the availability for that date
						foreach ( bool avail in group.Availability.AvailabilityForDate( date ) )
						{
							row.Append( ConstructCell( "", CellValues.String, (uint)( avail ? 2 :  3 ) ) );
						}

						sheetData.AppendChild( row );
					}

					worksheetPart.Worksheet.Save();
				}
			}
		}

		private static Stylesheet GenerateStylesheet()
		{
			Fonts fonts = new Fonts(
				new Font( // Index 0 - default
					new FontSize() { Val = 11 }

				),
				new Font( // Index 1 - header
					new FontSize() { Val = 11 },
					new Bold()
//					new Color() { Rgb = "FFFFFF" }

				) );

			Fills fills = new Fills(
					new Fill( new PatternFill() { PatternType = PatternValues.None } ), // Index 0 - default
					new Fill( new PatternFill() { PatternType = PatternValues.Gray125 } ), // Index 1 - default
					new Fill( new PatternFill( new ForegroundColor { Rgb = new HexBinaryValue( "FF92D050" ) } )
						{ PatternType = PatternValues.Solid } ), // Index 2 - available
					new Fill( new PatternFill( new ForegroundColor { Rgb = new HexBinaryValue( "FFFF0000" ) } ) 
						{ PatternType = PatternValues.Solid } ) // Index 3 - not available
				);

			Borders borders = new Borders(
					new Border(), // index 0 default
					new Border( // index 1 black border
						new LeftBorder( new Color() { Auto = true } ) { Style = BorderStyleValues.Thin },
						new RightBorder( new Color() { Auto = true } ) { Style = BorderStyleValues.Thin },
						new TopBorder( new Color() { Auto = true } ) { Style = BorderStyleValues.Thin },
						new BottomBorder( new Color() { Auto = true } ) { Style = BorderStyleValues.Thin },
						new DiagonalBorder() )
				);

			CellFormats cellFormats = new CellFormats(
					new CellFormat(), // default
					new CellFormat(
						new Alignment() { Horizontal = HorizontalAlignmentValues.Center,
							Vertical = VerticalAlignmentValues.Center } ) 
						{ FontId = 1, FillId = 0 }, // Headers 
					new CellFormat { FontId = 0, FillId = 2, BorderId = 1, ApplyBorder = true, ApplyFill = true }, // avail
					new CellFormat { FontId = 0, FillId = 3, BorderId = 1, ApplyBorder = true, ApplyFill = true }, // not avail
					new CellFormat(
						new Alignment() { Horizontal = HorizontalAlignmentValues.Right,
							Vertical = VerticalAlignmentValues.Center } ) // Dates
				);

			return new Stylesheet( fonts, fills, borders, cellFormats );
		}

		private static Cell ConstructCell( string value, CellValues dataType, uint styleIndex = 0 )
		{
			return new Cell() {
				CellValue = new CellValue( value ),
				DataType = new EnumValue<CellValues>( dataType ),
				StyleIndex = styleIndex
			};
		}

		/// <summary>
		/// Name and identity of each hostel
		/// </summary>
		private class Hostel
		{
			public Hostel( string hostelName, int hostelId )
			{
				Name = hostelName;
				Id = hostelId;
			}

			public string Name { get; set; }
			public int Id { get; set; }
		}

		/// <summary>
		/// A named group of hostels
		/// </summary>
		private class HostelGroup
		{
			public HostelGroup( string groupName )
			{
				Hostels = new List<Hostel>();
				Availability = new AvailabilityData();
				Name = groupName;
			}

			public HostelGroup AddHostel( params Hostel[] hostelsToAdd )
			{
				Hostels.AddRange( hostelsToAdd );
				return this;
			}

			public List< Hostel > Hostels { get; }
			public string Name { get; }
			public AvailabilityData Availability { get; }
		}

		/// <summary>
		/// The month request string used in the HTTP request and month name
		/// </summary>
		private class MonthRequest
		{
			public MonthRequest( string requestString, string monthName )
			{
				Request = requestString;
				Name = monthName;
			}

			public string Request
			{
				get;
			}

			public string Name
			{
				get;
			}
		}

		/// <summary>
		/// The availability data for a group of hostels over a number of days
		/// </summary>
		private class AvailabilityData
		{
			public AvailabilityData()
			{
			}

			/// <summary>
			/// Add an availability flag to the collection held for the date
			/// </summary>
			/// <param name="date"></param>
			/// <param name="available"></param>
			public void AddAvailability( string date, bool available, int hostelCount )
			{
				if ( dateFlags.ContainsKey( date ) == false )
				{
					dateFlags.Add( date, new List<bool>() );
					datesInOrder.Add( date );
				}

				if ( dateFlags[ date ].Count == hostelCount )
				{
					dateFlags[ date ].Add( available );
				}
			}

			public List<bool> AvailabilityForDate( string dateString )
			{
				return dateFlags[ dateString ];
			}

			public List<string> Dates
			{
				get
				{
					return datesInOrder;
				}
			}

			/// <summary>
			/// Availability flags (one per hostel) associated with dates
			/// </summary>
			private Dictionary< string, List<bool> > dateFlags = new Dictionary<string, List<bool>>();

			/// <summary>
			/// The date string in date order
			/// </summary>
			private List<string> datesInOrder = new List<string>();

		}

		private static readonly HttpClient client = new HttpClient();
	}
}
