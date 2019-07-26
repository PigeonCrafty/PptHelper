using Microsoft.Office.Core;

//using WORD = Microsoft.Office.Interop.Word;

namespace PptHelper.Languages
{
    public sealed class JAJP : LocalLanguage
    {
        public override string Tag { get; } = "ja-JP";
        public override string Name { get; } = "Japanese";
        public override string Location { get; } = "Japan";
        public override int Id { get; } = 0x0411;
        public override MsoLanguageID PptID { get; } = MsoLanguageID.msoLanguageIDJapanese;
        //public override WORD.WdLanguageID wdID { get; } = WORD.WdLanguageID.wdJapanese;

        public override bool IsFarEast { get; } = true;
        public override bool IsRightToLeft { get; } = false;
        public override bool IsSpecialFont { get; } = true;
        public override bool IsItalicToBold { get; } = true;
        public override string SpecialFont { get; } = "Meiryo"; // for AWS Japanese font
        public override string CurrencySymbolLocal1 { get; } = "¥";
        public override string CurrencySymbolLocal2 { get; } = "[$¥-ja-JP]";
        public override string CurrencySymbolLocal3 { get; } = "[$¥-411]";

        public override string[] CurrencyFormatLocal { get; } =
        {
            "\"¥\"#,##0.00;\"¥\"-#,##0.00",
            "\"¥\"#,##0.00;[Red]\"¥\"#,##0.00",
            "\"¥\"#,##0.00_);(\"¥\"#,##0.00)",
            "\"¥\"#,##0.00_);[Red](\"¥\"#,##0.00)"
        };

        public override string[] DateFormatLocal { get; } =
        {
            "yyyy/m/d;@",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "m/d;@",
            "m/d/yy;@",
            "mm/dd/yy;@",
            "m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy\"年\"m\"月\";@",
            "yyyy\"年\"m\"月\";@",
            "yyyy\"年\"m\"月\"d\"日\";@",
            "yyyy/m/d h:mm;@",
            "yyyy/m/d h:mm;@",
            "yyyy\"年\"m\"月\";@",
            "yyyy\"年\"m\"月\";@",
            "yyyy/m/d;@",
            "yyyy\"年\"m\"月\"d\"日\";@"
        };

        public override string GetLocAccounting(string numberFormat)
        {
            return "_ \"¥\"* #,##0.00_ ;_ \"¥\"* -#,##0.00_ ;_ \"¥\"* \"-\"??_ ;_ @_ ";
        }

        public override string GetLocFont(string sourceFont)
        {
            if (IsSpecialFont)
                return SpecialFont;
            return sourceFont;
        }
    }
}