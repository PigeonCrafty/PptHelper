using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using PptHelper.Languages;

namespace PptHelper
{
    public static class Common
    {
        #region Public Methods

        public static void WriteLine(string str)
        {
            //Console.WriteLine(str);
            log.Add(str);
            logSgl.Add(str);
        }

        public static void WriteAllText(string dirInput)
        {
            try
            {
                File.WriteAllText(dirInput + "\\log.txt", DateTime.Now + "\r\n" + "\r\n", Encoding.UTF8);
                foreach (var i in Common.log) File.AppendAllText(dirInput + "\\log.txt", i + "\r\n", Encoding.UTF8);
            }
            catch (Exception e)
            {
                Common.WriteLine("Fail to create log file! " + e.Message);
            }
        }

        public static void WriteSglText(string filePath)
        {
            try
            {
                var fileDirInfo = new DirectoryInfo(filePath);
                if (fileDirInfo.Parent == null) return;
                //string fileName = fileDirInfo.Name;
                //string langName = fileDirInfo.Parent.Name;
                //string logPath = dirInput + langName + "\\" + fileName + ".txt";
                var logPath = filePath + ".txt";

                File.WriteAllText(logPath, DateTime.Now + "\r\n", Encoding.UTF8);
                foreach (var i in Common.logSgl) File.AppendAllText(logPath, i + "\r\n", Encoding.UTF8);

                Common.logSgl.Clear();
            }
            catch (Exception e)
            {
                Common.WriteLine("Fail to create single log file! " + e.Message);
            }
        }

        public static LocalLanguage GetTargetLang(string input)
        {
            if (string.IsNullOrEmpty(input))
                throw new Exception("Empty or invalid language code! For example:\"en-US\"");

            switch (input.ToLower())
            {
                case "ja-jp":
                case "japanese":
                case "japan":
                case "jap":
                    return new JAJP();
                default: return new ENUS();
            }
        }
        /*
        public static string GetLangName(string input)
        {
            var fileDirInfo = new DirectoryInfo(input);
            if (fileDirInfo.Parent == null) return new Exception("Emply language Folder!").Message;
            var langName = fileDirInfo.Parent.Name;
            return langName;
        }

        public static LocalLanguage GetLangObj(string langName)
        {
            if (string.IsNullOrEmpty(langName))
                throw new Exception("Empty or invalid language code! For example:\"en-US\"");

            switch (langName.ToLower())
            {
                case "ar-sa":
                    return new ARSA();
                case "bg-bg":
                    return new BGBG();
                case "cs-cz":
                    return new CSCZ();
                case "da-dk":
                    return new CSCZ();
                case "de-de":
                    return new DEDE();
                case "el-gr":
                    return new ELGR();
                case "en-gb":
                    return new ENGB();
                case "es-es":
                    return new ESES();
                case "es-mx":
                    return new ESES();
                case "et-ee":
                    return new ETEE();
                case "fi-fi":
                    return new FIFI();
                case "fr-fr":
                    return new FRFR();
                case "he-il":
                    return new HEIL();
                case "hi-in":
                    return new HIIN();
                case "hr-hr":
                    return new HRHR();
                case "hu-hu":
                    return new HUHU();
                case "id-id":
                    return new IDID();
                case "it-it":
                    return new ITIT();
                case "ja-jp":
                    return new JAJP();
                case "ko-kr":
                    return new KOKR();
                case "lt-lt":
                    return new LTLT();
                case "lv-lv":
                    return new LVLV();
                case "nb-no":
                    return new NBNO();
                case "nl-nl":
                    return new NLNL();
                case "pl-pl":
                    return new PLPL();
                case "pt-br":
                    return new PTBR();
                case "pt-pt":
                    return new PTPT();
                case "ro-ro":
                    return new RORO();
                case "ru-ru":
                    return new RORO();
                case "sk-sk":
                    return new SKSK();
                case "sl-si":
                    return new SLSI();
                case "sr-latn-rs":
                    return new SRRS();
                case "sv-se":
                    return new SVSE();
                case "th-th":
                    return new THTH();
                case "tr-tr":
                    return new TRTR();
                case "uk-ua":
                    return new UKUA();
                case "vi-vn":
                    return new VIVN();
                case "zh-cn":
                    return new ZHCN();
                case "zh-tw":
                    return new ZHTW();
                default: return new ENUS();
            }
        }
        */
        #endregion

        #region Public Lists

        public static List<string> log = new List<string>();

        public static List<string> logSgl = new List<string>();

        #endregion

        #region Public Arrays

        public static string[] CurrencyFormatENU =
        {
            "$#,##0.00",
            "$#,##0.00;[Red]$#,##0.00",
            "$#,##0.00_);($#,##0.00)",
            "$#,##0.00_);[Red]($#,##0.00)"
        };

        public static string[] DateFormatENU =
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

        public static string[] DateFormatENU2 =
        {
            "m/d/yyyy",
            "[$-x-sysdate]dddd, mmmm dd, yyyy",
            "m/d;@",
            "m/d/yy;@",
            "mm/dd/yy;@",
            "[$-409]d-mmm;@",
            "[$-409]d-mmm-yy;@",
            "[$-409]dd-mmm-yy;@",
            "[$-409]mmm-yy;@",
            "[$-409]mmmm-yy;@",
            "[$-409]mmmm d, yyyy;@",
            "[$-409]m/d/yy h:mm AM/PM;@",
            "m/d/yy h:mm;@",
            "[$-409]mmmmm;@",
            "[$-409]mmmmm-yy;@",
            "m/d/yyyy;@",
            "[$-409]d-mmm-yyyy;@"
        };

        #endregion
    }
}