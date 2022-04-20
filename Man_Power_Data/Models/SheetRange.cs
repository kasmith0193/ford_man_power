using System.ComponentModel;

namespace Man_Power_Data.Models
{
	public class SheetRange : INotifyPropertyChanged
	{
		public event PropertyChangedEventHandler PropertyChanged;

		int _rowStart;
		public int RowStart
		{
			get => _rowStart;
			set
			{
				_rowStart = value < 1 ? 1 : value;
				Notify(nameof(RowStart));
			}
		}
		int _rowEnd;
		public int RowEnd
		{
			get => _rowEnd;
			set
			{
				_rowEnd = value < 1 ? 1 : value;
				Notify(nameof(RowEnd));
			}
		}

		int _columnStart;
		public int ColumnStart
		{
			get => _columnStart;
			set
			{
				_columnStart = value < 1 ? 1 : value;
				Notify(nameof(ColumnStart));
			}
		}

		int _columnEnd;
		public int ColumnEnd
		{
			get => _columnEnd;
			set
			{
				_columnEnd = value < 1 ? 1 : value;
				Notify(nameof(ColumnEnd));
			}
		}

		public SheetRange ()
		{
			RowStart = 1;
			RowEnd = 1;
			ColumnStart = 1;
			ColumnEnd = 1;
		}

		void Notify (string name)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}
	}
}
