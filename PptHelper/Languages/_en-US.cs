using Microsoft.Office.Core;

//using WORD = Microsoft.Office.Interop.Word;

namespace PptHelper.Languages
{
    public sealed class ENUS : LocalLanguage
    {
        public override string Tag { get; } = "en-US";
        public override string Name { get; } = "English";
        public override string Location { get; } = "United States";
        public override int Id { get; } = 0x0409;
        public override MsoLanguageID PptID { get; } = MsoLanguageID.msoLanguageIDEnglishUS;
        //public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdEnglishUS;

        public override bool IsFarEast { get; } = false;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = false;
        public override bool IsItalicToBold { get; } = false;
        public override string SpecialFont { get; } = null;
        public override string CurrencySymbolLocal1 { get; } = "\"$\"";
        public override string CurrencySymbolLocal2 { get; } = "[$$-en-US]";
        public override string CurrencySymbolLocal3 { get; } = "[$$-409]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "$#,##0.00",
            "$#,##0.00;[Red]$#,##0.00",
            "$#,##0.00_);($#,##0.00)",
            "$#,##0.00_);[Red]($#,##0.00)"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "m/d;@",
            "m/d/yy;@",
            "mm/dd/yy;@",
            "[$-en-US]d-mmm;@",
            "[$-en-US]d-mmm-yy;@",
            "[$-en-US]dd-mmm-yy;@",
            "[$-en-US]mmm-yy;@",
            "[$-en-US]mmmm-yy;@",
            "[$-en-US]mmmm d, yyyy;@",
            "[$-en-US]m/d/yy h:mm AM/PM;@",
            "m/d/yy h:mm;@",
            "[$-en-US]mmmmm;@",
            "[$-en-US]mmmmm-yy;@",
            "m/d/yyyy;@",
            "[$-en-US] d-mmm-yyyy;@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return numberFormat;
        }

        public override string GetLocFont(string sourceFont)
        {
            return sourceFont;
        }
    }
}