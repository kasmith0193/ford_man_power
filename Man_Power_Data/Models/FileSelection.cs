using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Man_Power_Data.Models
{
	public class FileSelection : INotifyPropertyChanged
	{
		public event PropertyChangedEventHandler PropertyChanged;

		string _fileLocation;
		public string FileLocation
		{
			get => _fileLocation;
			set
			{
				_fileLocation = value;
				Notify(nameof(FileLocation));
				Notify(nameof(FileName));
			}
		}

		public string FileName
		{
			get
			{
				try
				{
					return new Uri(FileLocation).Segments.Last();
				}
				catch (Exception)
				{
					return "";
				}
			}
		}

		public void SetFromDialog ()
		{
			// get new file location for the section
			var dialog = new OpenFileDialog()
			{
				DefaultExt = ".xlsx",
				Filter = "Excel Workbooks|*.xlsx",
				Title = "Select Source Sheet",
				CheckFileExists = true
			};

			if (FileLocation is string)
			{
				try
				{
					var file = new FileInfo(FileLocation);
					dialog.FileName = FileLocation;
					dialog.InitialDirectory = file.Directory.FullName;
				}
				catch (Exception) { }
			}

			if (dialog.ShowDialog() is true)
			{
				FileLocation = dialog.FileName;
			}
		}


		void Notify (string name)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}
	}
}
