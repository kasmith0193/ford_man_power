using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Man_Power_Data.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;


namespace Man_Power_Data
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	/// 

	public partial class MainWindow : Window
	{
		public const string OverviewComment = "These are the files that are being used to build the current configuration of the workbook. Changes from the tabs are represnted here.";
		public ReportConfiguration ReportConfiguration { get; set; } = ReportConfiguration.Default;
		public FileSelection TemplateFile { get; set; } = new FileSelection();
		public FileSelection DestinationFile { get; set; } = new FileSelection();


		private string[,] FullDepartmentFiles = new string[15, 4];


		string currentDropDownTitleString { get; set; }
		Section currentObject { get; set; }
		int currentObjectInt { get; set; }

		List<Dictionary<string, string>> allData { get; set; }


		public MainWindow()
		{
			InitializeComponent();
			//DataContext = this;
			allData = new List<Dictionary<string, string>>();
		}

		//-------------------------------------Spread Sheet methods---------------------------------------

		public void getSpreadsheetData(Section location)
		{
			try { 
				List<Dictionary<string, string>> allData = new List<Dictionary<string, string>>();

				var uniqueKeys = new Dictionary<string, bool>();

				using (SpreadsheetDocument spreadsheetdoc = SpreadsheetDocument.Open(location.FileSelection.FileLocation, false))
				{
					WorkbookPart workbookpart = spreadsheetdoc.WorkbookPart;
					IEnumerable<Sheet> sheets = spreadsheetdoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == location.Sheet);

					string relationshipId = sheets.First().Id.Value;
					WorksheetPart worksheetPart = (WorksheetPart)spreadsheetdoc.WorkbookPart.GetPartById(relationshipId);
					Worksheet workSheet = worksheetPart.Worksheet;
					SheetData sheetData = workSheet.GetFirstChild<SheetData>();
					IEnumerable<Row> rows = sheetData.Descendants<Row>();

					// load data into list of dictionaries
					foreach (Row row in rows.Where(x => x.RowIndex >= location.CopyFromRange.RowStart && x.RowIndex <= location.CopyFromRange.RowEnd))//lambda used to find rowindex between desired area
					{
						var dict = new Dictionary<string, string>();
						foreach (Cell cell in row.Descendants<Cell>().Where(x =>
						{//lambda used to find cells between had to convert ecxel alphanumberic code
							string val = x.CellReference;
							char[] trimval = { '1', '2', '3', '4', '5', '6', '7', '8', '9', '0' };
							val = val.Trim(trimval);
							int ci = 0;
							val = val.ToUpper();
							for (int ix = 0; ix < val.Length && val[ix] >= 'A'; ix++)
							{ ci = (ci * 26) + ((int)val[ix] - 64); }
							if (ci >= location.CopyFromRange.ColumnStart && ci <= location.CopyFromRange.ColumnEnd)
							{ return true; }
							else
							{ return false; }

						}))//Lambda ends here===========================================================
						{
							var col = Regex.Match(cell.CellReference.ToString(), @"[A-Za-z]+").Value;
							dict[col] = GetCellValue(spreadsheetdoc, cell);
							uniqueKeys[col] = true;
						}
						allData.Add(dict);
					}
				}

				// normalize the dictionaries to all have the same keys
				foreach (var key in uniqueKeys.Keys)
				{
					foreach (var row in allData)
					{
						if (!row.ContainsKey(key))
						{
							row[key] = "";
						}
					}
				}
				//add captured data to datagrid
				grid.Columns.Clear();
				var nodes = allData.FirstOrDefault();
				if (nodes?.Count > 0)
				{
					foreach (var node in nodes)
					{
						grid.Columns.Add(new DataGridTextColumn
						{
							Header = node.Key,
							Binding = new Binding($"[{node.Key}]")
						});
					}
				}
				grid.AutoGenerateColumns = false;
				grid.CanUserAddRows = false;
				grid.DataContext = allData;
			}
			catch(Exception exc)
			{
				MessageBox.Show("Exception caught. " + exc.Message);
			}
		}
		public static string GetCellValue(SpreadsheetDocument document, Cell cell)
		{
			SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
			string value = cell.CellValue?.InnerXml;

			if (value == null)
			{
				return "";
			}

			if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
			{
				return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
			}
			else
			{
				return value;
			}
		}

		//------------------------------------File List Methods------------------------------------------
		private void initiFileArray(string FileName)
		{
			for (int i = 0; i < 14; i++)
			{
				for (int j = 0; j < 4; j++)
				{
					FullDepartmentFiles[i, j] = FileName;
				}
			}
		}

		//------------------------------------Window Interaction Methods---------------
		async void MainWindow_Closing(object sender, CancelEventArgs e)
		{
			try
			{
				await ReportConfiguration.SaveSections();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ReportConfiguration.GetMessageForException(ex));
			}
		}


		async void MainWindow_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				await ReportConfiguration.LoadSections();
			}
			catch (FileNotFoundException)
			{
				// if the file isn't found, then we implicitly create a new configuration file
				//   so don't do anything here
			}
			catch (Exception ex)
			{
				MessageBox.Show(ReportConfiguration.GetMessageForException(ex));
			}
		}



		private void NewFile_Click(object sender, RoutedEventArgs e)
		{
			Saveoption();

			
			SaveFileDialog newfileDialog = new SaveFileDialog();
			newfileDialog.Filter = "JSON File (*.json)|*.json";
			if(newfileDialog.ShowDialog() == true)
			{
				if (newfileDialog.FileName != null)
				{
					ReportConfiguration = new ReportConfiguration();
					FileStream file = File.Create(newfileDialog.FileName);
					file.Close();

					ReportConfiguration.SectionConfigurationPath = newfileDialog.FileName;
					ReportConfiguration.Sections.Add(new Section());

					var binddata = new Binding() { Source = ReportConfiguration.Sections };
					OverViewDataGrid.SetBinding(DataGrid.ItemsSourceProperty, binddata);

					var bindlabel = new Binding() { Source = ReportConfiguration.SectionConfigurationPath };
					ReportConfigPathLabel.SetBinding(Label.ContentProperty, bindlabel);
				}
				
			}
		}
		async void SaveFile_Click(object sender, RoutedEventArgs e)
		{
			SaveFileDialog savefileDialog = new SaveFileDialog();
			savefileDialog.Filter = "JSON File (*.json)|*.json";
			try
			{
				if (savefileDialog.ShowDialog() == true)
				{
					await ReportConfiguration.SaveSections(savefileDialog.FileName);
					ReportConfiguration.SectionConfigurationPath = savefileDialog.FileName;
					await ReportConfiguration.LoadSections();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ReportConfiguration.GetMessageForException(ex));
			}
		}
		async void OpenFile_Click(object sender, RoutedEventArgs e)
		{


			Saveoption();

			try
			{
				var opendialog = new Microsoft.Win32.OpenFileDialog();
				opendialog.Filter = "JSON File (*.json)|*.json";
				if (opendialog.ShowDialog() == true)
				{
					ReportConfiguration.SectionConfigurationPath = opendialog.FileName;
					//loadstuff();
					await ReportConfiguration.LoadSections();
				}
			}
			catch (FileNotFoundException)
			{
				MessageBox.Show("Can not find this file.");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ReportConfiguration.GetMessageForException(ex));
			}
		}
		async void Saveoption()
		{
			if (MessageBoxResult.Yes == MessageBox.Show("Save current Config?", "Save?", MessageBoxButton.YesNo))
			{
				try
				{
					await ReportConfiguration.SaveSections();
					await ReportConfiguration.LoadSections();
					MessageBox.Show("File has been saved.");
				}
				catch (Exception ex)
				{
					MessageBox.Show(ReportConfiguration.GetMessageForException(ex));
				}
			}
		}

		private void GatherFromDialog_Click(object sender, RoutedEventArgs e)
		{
			var dialog = new Microsoft.Win32.OpenFileDialog();
			// dialog.DefaultExt = ".xls|.xlsx";
			//dialog.Filter = "Excel (.xls)|*.xls|(.xlsx)|*.xlsx";

			if (dialog.ShowDialog() == true)
			{
				if (dialog.FileName.Length > 0)
				{
					GatherDataFromText.Text = dialog.FileName;

					currentObject.FileSelection.FileLocation = dialog.FileName;
				}
			}
		}//done
		private void GatherFromText_Changed(object sender, RoutedEventArgs e)
		{
			ReportConfiguration.Sections[currentObjectInt].FileSelection.FileLocation = GatherDataFromText.Text;
		}//done
		private void DepartmentSelect_Change(Object sender, SelectionChangedEventArgs e)
		{
			try
			{
				currentObject = (Section)DepartmentSelect.SelectedItem;
				OverViewDataGrid.Items.Refresh();
				GatherDataFromText.Text = currentObject.FileSelection.FileLocation;
				getSpreadsheetData(currentObject);
			}
			catch (Exception exc)
			{
				MessageBox.Show("Exception caught. " + exc.Message);
			}

		}//done
		private void GenerateReport_Click(object sender, RoutedEventArgs e)
		{
			//var location = new Section();
			//getSpreadsheetData(location);
			try{
				ReportGeneration reportGeneration = new ReportGeneration();
				reportGeneration.addSections(ReportConfiguration.Sections);
				reportGeneration.CreateReport(DestinationFile.FileLocation, TemplateFile.FileLocation);
				MessageBox.Show("Report Done");
			}
			catch(Exception exc)
			{
				MessageBox.Show("Exception caught. " + exc.Message);
			}
		
			
			

		}

		private void SectionDelete_Click (object sender, RoutedEventArgs e)
		{
			if (sender is Button button && button.DataContext is Section section)
			{
				ReportConfiguration.Sections.Remove(section);
			}
		}

		private void AddSection_Click (object sender, RoutedEventArgs e)
		{
			ReportConfiguration.Sections.Add(new Section());
		}

		private void SelectFile_Click (object sender, RoutedEventArgs e)
		{
			if (sender is FrameworkElement el && el.DataContext is FileSelection selection)
			{
				selection.SetFromDialog();
			}
		}
	}
}









