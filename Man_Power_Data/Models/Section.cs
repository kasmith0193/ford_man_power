using System;
using System.ComponentModel;
using System.Linq;

namespace Man_Power_Data.Models
{
	public class Section : INotifyPropertyChanged
	{
		public event PropertyChangedEventHandler PropertyChanged;

		string _title;
		public string Title
		{
			get => _title;
			set
			{
				_title = value;
				Notify(nameof(Title));
			}
		}

		string _sheet;
		public string Sheet
		{
			get => _sheet;
			set
			{
				_sheet = value;
				Notify(nameof(Sheet));
			}
		}

		FileSelection _fileSelection;
		public FileSelection FileSelection
		{
			get => _fileSelection;
			set
			{
				_fileSelection = value;
				Notify(nameof(FileSelection));
			}
		}

		SheetRange _copyFromRange;
		public SheetRange CopyFromRange
		{
			get => _copyFromRange;
			set
			{
				_copyFromRange = value;
				Notify(nameof(CopyFromRange));
			}
		}

		public Section ()
		{
			Title = "New Section";
			Sheet = "Sheet 1";
			FileSelection = new FileSelection();
			CopyFromRange = new SheetRange();
		}

		void Notify (string name)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}
	}
}
