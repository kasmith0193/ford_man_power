using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Man_Power_Data.Models
{
	public class ReportConfiguration : INotifyPropertyChanged
	{
		public event PropertyChangedEventHandler PropertyChanged;


		ObservableCollection<Section> _sections = new ObservableCollection<Section>();
		public ObservableCollection<Section> Sections
		{
			get => _sections;
			set
			{
				_sections = value;
				Notify(nameof(Sections));
			}
		}

		string _path;
		public string SectionConfigurationPath
		{
			get => _path;
			set
			{
				_path = value;
				Sections = new ObservableCollection<Section>();
				Notify(nameof(SectionConfigurationPath));
			}
		}		




		public async Task LoadSections ()
		{
			if (_path is string)
			{
				using var readStream = File.OpenRead(SectionConfigurationPath);
				Sections = await JsonSerializer.DeserializeAsync<ObservableCollection<Section>>(readStream);
			}
			else
			{
				throw new FileNotFoundException();
			}
		}

		public async Task SaveSections (string filename = null)
		{
			if (_path is string)
			{
				if (Sections.Count > 0)
				{
					using var writeStream = File.OpenWrite(filename ?? SectionConfigurationPath);
					await JsonSerializer.SerializeAsync(writeStream, Sections);
					writeStream.SetLength(writeStream.Position);
				}
				else
				{
					throw new InvalidOperationException();
				}
			}
			else
			{
				throw new FileNotFoundException();
			}
		}

		public string GetMessageForException (Exception ex)
		{
			return ex switch
			{
				FileNotFoundException => $"The file \"{SectionConfigurationPath}\" does not exist.",
				InvalidOperationException => "There is no section data to save.",
				UnauthorizedAccessException => $"You do not have permission to access the file \"{SectionConfigurationPath}\".",
				DirectoryNotFoundException => $"The directory containing \"{SectionConfigurationPath}\" does not exist.",
				JsonException => $"The section configuration stored in \"{SectionConfigurationPath}\" cannot be read.",
				PathTooLongException => $"The path \"{SectionConfigurationPath}\" is too long.",
				_ => "An unexpected error occurred."
			};
		}

		void Notify (string propertyName)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}

		public static ReportConfiguration Default => new ReportConfiguration { _path = "LocationOverview.json" };
	}
}
