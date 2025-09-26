using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace Excel
{
    public sealed class ProductSummaryConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length < 2)
            {
                return string.Empty;
            }

            if (values[0] is not string productName || string.IsNullOrWhiteSpace(productName))
            {
                return string.Empty;
            }

            var dictionary = ResolveDictionary(values[1]);
            if (dictionary == null)
            {
                return string.Empty;
            }

            if (!dictionary.TryGetValue(productName, out var summary))
            {
                return string.Empty;
            }

            return summary.BuildDisplayText(culture);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        private static IReadOnlyDictionary<string, ProductSummary>? ResolveDictionary(object value)
        {
            if (value is IReadOnlyDictionary<string, ProductSummary> typed)
            {
                return typed;
            }

            if (value is IDictionary dictionary)
            {
                var result = new Dictionary<string, ProductSummary>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry entry in dictionary)
                {
                    if (entry.Key is string key && entry.Value is ProductSummary summary)
                    {
                        result[key] = summary;
                    }
                }

                return result;
            }

            return null;
        }
    }
}
