using Microsoft.Office.Core;

//using WORD = Microsoft.Office.Interop.Word;

namespace PptHelper.Languages
{
    public abstract class LocalLanguage
    {
        public abstract string Tag { get; }
        public abstract string Name { get; }
        public abstract string Location { get; }
        public abstract int Id { get; }
        public abstract MsoLanguageID PptID { get; }
        //public abstract WORD.WdLanguageID wdID { get; }

        public abstract bool IsFarEast { get; }
        public abstract bool IsRightToLeft { get; }
        public abstract bool IsSpecialFont { get; }
        public abstract bool IsItalicToBold { get; }
        public abstract string SpecialFont { get; }
        public abstract string CurrencySymbolLocal1 { get; }
        public abstract string CurrencySymbolLocal2 { get; }
        public abstract string CurrencySymbolLocal3 { get; }

        public abstract string[] CurrencyFormatLocal { get; }
        public abstract string[] DateFormatLocal { get; }

        public abstract string GetLocFont(string sourceFont);
        public abstract string GetLocAccounting(string numberFormat);

        public string GetLocCurrency(string numberFormat)
        {
            var type = GetCurrencyType(numberFormat);
            var locFormat = CurrencyFormatLocal[type];
            return locFormat;
        }

        public int GetCurrencyType(string numFormat)
        {
            if (!numFormat.Contains(";")) return 0;

            if (!numFormat.Contains("[Red]") && !numFormat.Contains("_)"))
                return 0;
            if (numFormat.Contains("[Red]") && !numFormat.Contains("_)"))
                return 1;
            if (!numFormat.Contains("[Red]") && numFormat.Contains("_)"))
                return 2;
            if (numFormat.Contains("[Red]") && numFormat.Contains("_)")) return 3;

            return 0;
        }

        public string GetLocDate(string numberFormat)
        {
            for (var i = 0; i < DateFormatLocal.Length; i++)
                if (Common.DateFormatENU[i].Contains(numberFormat) || Common.DateFormatENU2[i].Contains(numberFormat))
                    return DateFormatLocal[i];
            return DateFormatLocal[0];
        }
    }
}