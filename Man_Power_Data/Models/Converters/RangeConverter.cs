using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace Man_Power_Data.Models.Converters
{
	public class RangeStringConverter : IValueConverter
	{
		public object Convert (object value, Type targetType, object parameter, CultureInfo culture)
		{
			if (targetType == typeof(string) && value is SheetRange range)
			{
				return ToString(range);
			}
			else if (targetType == typeof(SheetRange) && value is string str)
			{
				return FromString(str);
			}
			else
			{
				return DependencyProperty.UnsetValue;
			}
		}

		public object ConvertBack (object value, Type targetType, object parameter, CultureInfo culture)
		{
			return Convert(value, targetType, parameter, culture);
		}

		private static object ToString (SheetRange range)
		{
			return $"{IntToAlpha(range.ColumnStart)}{range.RowStart}:{IntToAlpha(range.ColumnEnd)}{range.RowEnd}";
		}


		private static readonly Regex RangeRegex = new Regex(@"([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)");
		private static object FromString (string range)
		{
			var match = RangeRegex.Match(range);
			if (!match.Success)
			{
				return DependencyProperty.UnsetValue;
			}

			int rowStart = int.Parse(match.Groups[2].Value);
			int rowEnd = int.Parse(match.Groups[4].Value);

			if (rowStart > rowEnd)
			{
				// start must be less than end
				int temp = rowStart;
				rowStart = rowEnd;
				rowEnd = temp;
			}

			string strColStart = match.Groups[1].Value;
			string strColEnd = match.Groups[3].Value;

			int colStart = AlphaToInt(strColStart);
			int colEnd = AlphaToInt(strColEnd);

			if (colStart > colEnd)
			{
				// start must be less than end
				int temp = colStart;
				colStart = colEnd;
				colEnd = temp;
			}

			return new SheetRange()
			{
				RowStart = rowStart,
				RowEnd = rowEnd,
				ColumnStart = colStart,
				ColumnEnd = colEnd
			};
		}

		private static int AlphaToInt (string alpha)
		{
			alpha = alpha.ToUpper();
			int sum = 0;
			for (int i = alpha.Length; i > 0; i--)
			{
				int power = (int)Math.Pow(26, i - 1);
				int digit = alpha[alpha.Length - i] - 'A' + 1;
				sum += digit * power;
			}
			return sum;
		}

		private static string IntToAlpha (int n)
		{
			var alpha = new StringBuilder();
			do
			{
				int rem = n % 26;
				alpha.Insert(0, (char)(rem + 'A' - 1));
				n /= 26;
			} while (n > 0);
			return alpha.ToString();
		}
	}
}
