using Microsoft.Office.Core;
using System;
using System.IO;
using System.Reflection;
using PptHelper;
using PptHelper.Languages;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;

namespace PptHelper
{
    public partial class PowerPointHandler
    {
        #region Constructor

        public PowerPointHandler(string filePath, string targetLang)
        {
            app = new POWERPOINT.Application { DisplayAlerts = POWERPOINT.PpAlertLevel.ppAlertsNone };
            LocalLang = Common.GetTargetLang(targetLang);
            //LangName = Common.GetLangName(filePath);
            //LocalLang = Common.GetLangObj(LangName.ToLower());
        }

        #endregion

        public void PptMain(string filePath)
        {
            try
            {
                pre = app.Presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse,
                    MsoTriState.msoFalse);
            }
            catch (Exception)
            {
                Common.WriteLine("[Failed to Open]: " + filePath);
                app.Quit();
                return;
            }

            var pptFileInfo = new FileInfo(filePath);

            // 1. Process Normal Slides
            var slideNum = 0;
            foreach (POWERPOINT.Slide slide in pre.Slides)
            {
                slideNum += 1;
                var shapeNum = 0;

                foreach (POWERPOINT.Shape shape in slide.Shapes)
                {
                    shapeNum += 1;

                    // if the shape is text
                    if (shape.HasTextFrame == MsoTriState.msoTrue || shape.Type == MsoShapeType.msoTextBox)
                    {
                        try
                        {
                            ProcText(shape);
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("<!> Error in Normal Slide " + slideNum + " Shape " + shapeNum);
                        }
                    }
                    // if the shape is table
                    else if (shape.HasTable == MsoTriState.msoTrue || shape.Type == MsoShapeType.msoTable)
                    {
                        try
                        {
                            ProcTable(shape);
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("<!> Error in Normal Slide " + slideNum + " Shape " + shapeNum);
                        }
                    }
                    // if the shape is chart
                    else if (shape.HasChart == MsoTriState.msoTrue || shape.Type == MsoShapeType.msoChart)
                    {
                        try
                        {
                            ProcChart(shape);
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("<!> Error in Normal Slide " + slideNum + " Shape " + shapeNum);
                        }
                    }
                    // if the shape is smart art
                    else if (shape.HasSmartArt == MsoTriState.msoTrue || shape.Type == MsoShapeType.msoSmartArt)
                    {
                        try
                        {
                            ProcSmartArt(shape);
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("<!> Error in Normal Slide " + slideNum + " Shape " + shapeNum);
                        }
                    }
                    // if the shape is group
                    else if (shape.Type == MsoShapeType.msoGroup)
                    {
                        try
                        {
                            ProcGroups(shape);
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("<!> Error in Normal Slide " + slideNum + " Shape " + shapeNum);
                        }
                    }
                    // if the shape is PlaceHolder
                    else if (shape.Type == MsoShapeType.msoPlaceholder)
                    {
                        try
                        {
                            ProcPlaceHolder(shape); // Not working yet
                            Common.WriteLine(">> Normal Slide " + slideNum + " Shape " + shapeNum);
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("<!> Error in Normal Slide " + slideNum + " Shape " + shapeNum);
                        }
                    }
                    // if the shape is only picture or other type, do nothing but report the type of shape
                    else
                    {
                        var shapeType = shape.Type.ToString();
                        if (shape.Type != MsoShapeType.msoPicture && shape.Type != MsoShapeType.msoLine
                        ) // Report all untouched shapes except picture and line
                        {
                            Common.WriteLine("Didn't update " + shapeType + ":");
                            Common.WriteLine(">> Normal Slide " + slideNum + " Shape " + shapeNum);
                        }
                    }
                }
            }

            var slideNum2 = 0;
            foreach (POWERPOINT.Slide slide in pre.Slides) // each slide
            {
                slideNum += 1;
                var noteNum = 0;
                if (slide.HasNotesPage == MsoTriState.msoTrue) // only if has Notes Page
                    for (var i = 1; i <= slide.NotesPage.Count; i++) // each note page
                    {
                        noteNum += 1;
                        foreach (POWERPOINT.Shape shape in slide.NotesPage[i].Shapes) // each shape
                            if (shape.HasTextFrame == MsoTriState.msoTrue || shape.Type == MsoShapeType.msoTextBox)
                                try
                                {
                                    ProcText(shape);
                                }
                                catch (Exception)
                                {
                                    Common.WriteLine("<!> Error in Note Page " + slideNum2 + " Shape " + noteNum);
                                }
                    }
            }

            // 2. Process Slide Master
            var masterShapeNum = 0;
            foreach (POWERPOINT.Shape shape in pre.SlideMaster.Shapes)
            {
                masterShapeNum += 1;
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    try
                    {
                        ProcText(shape);
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("<!> Error in Slide Master Shape " + masterShapeNum);
                    }
                }
                else if (shape.Type == MsoShapeType.msoGroup)
                {
                    try
                    {
                        ProcGroups(shape);
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("<!> Error in Slide Master Shape " + masterShapeNum);
                    }
                }
                else
                {
                    var shapeType = shape.Type.ToString();
                    if (shape.Type != MsoShapeType.msoPicture && shape.Type != MsoShapeType.msoLine
                    ) // Report all untouched shapes except pictures
                    {
                        Common.WriteLine("Didn't update " + shapeType + ":");
                        Common.WriteLine(">> SlideMaster Shape " + masterShapeNum);
                    }
                }
            }

            // 3. Process Slide Master - Customer Layouts
            var custLayoutSlideNum = 0;
            foreach (POWERPOINT.CustomLayout customLayout in pre.SlideMaster.CustomLayouts)
            {
                custLayoutSlideNum += 1;
                var custLayoutShapeNum = 0;
                foreach (POWERPOINT.Shape shape in customLayout.Shapes)
                {
                    custLayoutShapeNum += 1;

                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        try
                        {
                            ProcText(shape);
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("<!> Error in Slide Master Custom Layout TextFrame");
                        }
                    }
                    else if (shape.Type == MsoShapeType.msoGroup)
                    {
                        try
                        {
                            ProcGroups(shape);
                        }
                        catch (Exception)
                        {
                            Common.WriteLine("<!> Error in Slide Master Custom Layout TextFrame");
                        }
                    }
                    else
                    {
                        var shapeType = shape.Type.ToString();
                        if (shape.Type != MsoShapeType.msoPicture && shape.Type != MsoShapeType.msoLine
                        ) // Report all untouched shapes except picture and line
                        {
                            Common.WriteLine("Didn't update " + shapeType + ":");
                            Common.WriteLine(">> Customer Layout Slide " + custLayoutSlideNum + " Shape " +
                                             custLayoutShapeNum);
                        }
                    }
                }
            }

            // 4. Change Handout Master
            //TODO
            var handoutMasterShapeNum = 0;
            foreach (POWERPOINT.Shape shape in pre.HandoutMaster.Shapes)
            {
                handoutMasterShapeNum += 1;
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    try
                    {
                        ProcText(shape);
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("<!> Error in HandoutMaster Shape: " + handoutMasterShapeNum);
                    }
                }
                else
                {
                    var shapeType = shape.Type.ToString();
                    if (shape.Type != MsoShapeType.msoPicture && shape.Type != MsoShapeType.msoLine
                    ) // Report all untouched shapes except picture and line
                    {
                        Common.WriteLine("Didn't update " + shapeType + ":");
                        Common.WriteLine(">> HandoutMaster Shape " + handoutMasterShapeNum);
                    }
                }
            }

            // 5. Change Note Master
            //TODO delete
            var noteMasterShapeNum = 0;
            foreach (POWERPOINT.Shape shape in pre.NotesMaster.Shapes)
            {
                noteMasterShapeNum += 1;
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    try
                    {
                        ProcText(shape);
                        //if (noteMasterShapeNum == 2)
                        //{
                        //    shape.TextFrame2.TextRange.Select();
                        //    // TODO
                        //}
                    }
                    catch (Exception)
                    {
                        Common.WriteLine("<!> Error in oteMaster Shape " + noteMasterShapeNum);
                    }
                }
                else
                {
                    var shapeType = shape.Type.ToString();
                    if (shape.Type != MsoShapeType.msoPicture && shape.Type != MsoShapeType.msoLine
                    ) // Report all untouched shapes except picture and line
                    {
                        Common.WriteLine("Didn't update " + shapeType + ":");
                        Common.WriteLine(">> NoteMaster Shape " + noteMasterShapeNum);
                    }
                }
            }

            // Save and close
            try
            {
                // 6. Prompt complete info when each file processed
                //Console.WriteLine(LangName + "\\" + pre.Name + " Complete!" + "\n");
                LangName = "";
                Dispose();
            }
            catch (Exception e)
            {
                // Save as another file with _New_
                var newName = pptFileInfo.Name + "_New_" + pptFileInfo.Extension;
                var format = POWERPOINT.PpSaveAsFileType.ppSaveAsDefault;
                pre.SaveAs(newName, format, MsoTriState.msoFalse);
                Common.WriteLine("Error occurs during saving: " + e.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        #region Private Fields

        private POWERPOINT.Application app { get; }

        private POWERPOINT.Presentation pre { get; set; }

        private string LangName { get; set; }

        private LocalLanguage LocalLang { get; }

        private readonly Missing misValue = Missing.Value;

        #endregion
    }
}