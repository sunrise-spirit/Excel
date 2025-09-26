using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    public sealed class ProductSummary
    {
        public ProductSummary(string product, decimal oldAmount, decimal newAmount)
        {
            Product = product;
            OldAmount = DecimalHelper.Round2(oldAmount);
            NewAmount = DecimalHelper.Round2(newAmount);
            AmountDelta = DecimalHelper.Round2(NewAmount - OldAmount);
            AmountDeltaPercent = DecimalHelper.CalculatePercentChange(NewAmount, OldAmount);
        }

        public string Product { get; }
        public decimal OldAmount { get; }
        public decimal NewAmount { get; }
        public decimal AmountDelta { get; }
        public decimal? AmountDeltaPercent { get; }

        public bool ShouldDisplay(decimal thresholdPercent)
        {
            if (AmountDeltaPercent.HasValue)
            {
                return Math.Abs(AmountDeltaPercent.Value) >= thresholdPercent;
            }

            return !DecimalHelper.IsEffectivelyZero(AmountDelta);
        }

        public string BuildDisplayText(CultureInfo culture)
        {
            var oldText = OldAmount.ToString("N2", culture);
            var newText = NewAmount.ToString("N2", culture);
            var deltaText = AmountDelta.ToString("N2", culture);
            var percentText = AmountDeltaPercent.HasValue
                ? AmountDeltaPercent.Value.ToString("N2", culture) + "%"
                : "—";

            return $"Σ: {oldText} → {newText} (Δ {deltaText}, {percentText})";
        }
    }
}
