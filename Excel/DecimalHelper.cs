using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    internal static class DecimalHelper
    {
        private const int Precision = 2;

        public static decimal Round2(decimal value) => Math.Round(value, Precision, MidpointRounding.AwayFromZero);

        public static decimal? Round2(decimal? value) => value.HasValue ? Round2(value.Value) : (decimal?)null;

        public static decimal? CalculatePercentChange(decimal? newValue, decimal? oldValue)
        {
            if (newValue.HasValue && oldValue.HasValue)
            {
                var newAmount = newValue.Value;
                var oldAmount = oldValue.Value;

                if (IsEffectivelyZero(oldAmount))
                {
                    return IsEffectivelyZero(newAmount) ? 0m : 100m;
                }

                return Round2((newAmount - oldAmount) / oldAmount * 100m);
            }

            if (newValue.HasValue && (!oldValue.HasValue || IsEffectivelyZero(oldValue.Value)))
            {
                return IsEffectivelyZero(newValue.Value) ? 0m : 100m;
            }

            if (!newValue.HasValue && oldValue.HasValue)
            {
                return IsEffectivelyZero(oldValue.Value) ? 0m : -100m;
            }

            return null;
        }

        public static bool IsEffectivelyZero(decimal value) => Math.Abs(value) < 0.0001m;
    }
}
