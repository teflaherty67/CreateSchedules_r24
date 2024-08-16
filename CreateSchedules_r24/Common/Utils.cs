﻿using Autodesk.Revit.DB;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace CreateSchedules_r24
{
    internal static class Utils
    {

        #region Areas Scheme

        internal static AreaScheme GetAreaSchemeByName(Document doc, string schemeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(AreaScheme));

            foreach (AreaScheme areaScheme in collector)
            {
                if (areaScheme.Name == schemeName)
                {
                    return areaScheme;
                }
            }

            return null;
        }

        #endregion

        #region Design Options

        internal static List<DesignOption> getAllDesignOptions(Document curDoc)
        {
            FilteredElementCollector curCol = new FilteredElementCollector(curDoc);
            curCol.OfCategory(BuiltInCategory.OST_DesignOptions);

            List<DesignOption> doList = new List<DesignOption>();
            foreach (DesignOption curOpt in curCol)
            {
                doList.Add(curOpt);
            }

            return doList;
        }

        internal static DesignOption getDesignOptionByName(Document curDoc, string designOpt)
        {
            //get all design options
            List<DesignOption> doList = getAllDesignOptions(curDoc);

            foreach (DesignOption curOpt in doList)
            {
                if (curOpt.Name == designOpt)
                {
                    return curOpt;
                }
            }

            return null;
        }

        #endregion

        #region Levels

        internal static List<ElementId> GetAllLevelIds(Document doc)
        {
            FilteredElementCollector colLevelId = new FilteredElementCollector(doc);
            colLevelId.OfClass(typeof(Level));

            List<Level> sortedList = colLevelId.Cast<Level>().ToList().OrderBy(x => x.Elevation).ToList();

            List<ElementId> returnList = new List<ElementId>();

            foreach (Level curLevel in sortedList)
            {
                if (curLevel.Name.Contains("Floor") || curLevel.Name.Contains("Level"))
                    returnList.Add(curLevel.Id);
            }

            return returnList;
        }

        internal static List<Level> GetLevelByNameContains(Document doc, string levelWord)
        {
            List<Level> levels = GetAllLevels(doc);

            List<Level> returnList = new List<Level>();

            foreach (Level curLevel in levels)
            {
                if (curLevel.Name.Contains(levelWord))
                    returnList.Add(curLevel);
            }

            return returnList;
        }

        internal static Level GetLevelByName(Document doc, string levelName)
        {
            List<Level> levels = GetAllLevels(doc);

            foreach (Level curLevel in levels)
            {
                Debug.Print(curLevel.Name);

                if (curLevel.Name.Equals(levelName))
                    return curLevel;
            }

            return null;
        }

        public static List<Level> GetAllLevels(Document doc)
        {
            FilteredElementCollector colLevels = new FilteredElementCollector(doc);
            colLevels.OfCategory(BuiltInCategory.OST_Levels);

            List<Level> levels = new List<Level>();
            foreach (Element x in colLevels.ToElements())
            {
                if (x.GetType() == typeof(Level))
                {
                    levels.Add((Level)x);
                }
            }

            return levels;
            //order list by elevation
            //m_levels = (From l In m_levels Order By l.Elevation).tolist()
        }

        #endregion

        #region Parameters

        internal static List<Parameter> GetParametersByName(Document doc, List<string> paramNames, BuiltInCategory cat)
        {
            List<Parameter> m_returnList = new List<Parameter>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(cat);

            foreach (string curName in paramNames)
            {
                Parameter curParam = collector.FirstElement().LookupParameter(curName);

                if (curParam != null)
                    m_returnList.Add(curParam);
            }

            return m_returnList;
        }

        internal static ElementId GetProjectParameterId(Document doc, string name)
        {
            ParameterElement pElem = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterElement))
                .Cast<ParameterElement>()
                .Where(e => e.Name.Equals(name))
                .FirstOrDefault();

            return pElem?.Id;
        }

        internal static ElementId GetBuiltInParameterId(Document doc, BuiltInCategory cat, BuiltInParameter bip)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(cat);

            Parameter curParam = collector.FirstElement().get_Parameter(bip);

            return curParam?.Id;
        }

        internal static Parameter GetParameterByName(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName)
                    return curParam;
            }

            return null;
        }

        internal static bool SetParameterValue(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);

            if (curParam != null)
            {
                curParam.Set(value);
                return true;
            }

            return false;
        }

        #endregion

        #region Ribbon

        internal static BitmapImage BitmapToImageSource(Bitmap bm)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }

        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel = GetRibbonPanelByName(app, tabName, panelName);

            if (curPanel == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            return curPanel;
        }

        private static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tempPanel in app.GetRibbonPanels(tabName))
            {
                if (tempPanel.Name == panelName)
                    return tempPanel;
            }

            return null;
        }

        #endregion

        #region Schedules

        internal static ViewSchedule GetScheduleByNameContains(Document doc, string scheduleString)
        {
            List<ViewSchedule> m_scheduleList = GetAllSchedules(doc);

            foreach (ViewSchedule curSchedule in m_scheduleList)
            {
                if (curSchedule.Name.Contains(scheduleString))
                    return curSchedule;
            }

            return null;
        }

        private static List<ViewSchedule> GetAllSchedules(Document doc)
        {
            {
                List<ViewSchedule> m_schedList = new List<ViewSchedule>();

                FilteredElementCollector curCollector = new FilteredElementCollector(doc);
                curCollector.OfClass(typeof(ViewSchedule));

                //loop through views and check if schedule - if so then put into schedule list
                foreach (ViewSchedule curView in curCollector)
                {
                    if (curView.ViewType == ViewType.Schedule)
                    {
                        m_schedList.Add((ViewSchedule)curView);
                    }
                }

                return m_schedList;
            }
        }

        internal static List<ViewSchedule> GetAllScheduleByNameContains(Document doc, string schedName)
        {
            List<ViewSchedule> m_scheduleList = GetAllSchedules(doc);

            List<ViewSchedule> m_returnList = new List<ViewSchedule>();

            foreach (ViewSchedule curSchedule in m_scheduleList)
            {
                if (curSchedule.Name.Contains(schedName))
                    m_returnList.Add(curSchedule);
            }

            return m_returnList;
        }

        internal static ViewSchedule CreateAreaSchedule(Document doc, string schedName, AreaScheme curAreaScheme)
        {
            ElementId catId = new ElementId(BuiltInCategory.OST_Areas);
            ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId, curAreaScheme.Id);
            newSchedule.Name = schedName;

            return newSchedule;
        }

        internal static ViewSchedule CreateSchedule(Document doc, BuiltInCategory curCat, string name)
        {
            ElementId catId = new ElementId(curCat);
            ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId);
            newSchedule.Name = name;

            return newSchedule;
        }

        internal static void AddFieldsToSchedule(Document doc, ViewSchedule newSched, List<Parameter> paramList)
        {
            foreach (Parameter curParam in paramList)
            {
                SchedulableField newField = new SchedulableField(ScheduleFieldType.Instance, curParam.Id);
                newSched.Definition.AddField(newField);
            }
        }

        internal static void DuplicateAndConfigureEquipmentSchedule(Document curDoc)
        {
            // duplicate the first schedule with "Roof Ventilation Equipment" in the name
            List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Roof Ventilation Equipment");
            ViewSchedule dupSched = listSched.FirstOrDefault();

            if (dupSched == null)
            {
                // call another method to create one

                return; // no schedule to duplicate
            }

            // duplicate the schedule
            ViewSchedule equipmentSched = curDoc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;

            // rename the duplicated schedule to the new elevation
            string originalName = equipmentSched.Name;
            string[] schedTitle = originalName.Split('-');

            equipmentSched.Name = schedTitle[0] + "- Elevation " + GlobalVars.ElevDesignation;

            Utils.SetParameterValue(equipmentSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

            //// set the design option to the specified elevation designation
            //DesignOption curOption = Utils.getDesignOptionByName(curDoc, "Elevation : " + Globals.ElevDesignation);

            //Parameter doParam = veneerSched.get_Parameter(BuiltInParameter.VIEWER_OPTION_VISIBILITY);

            //doParam.Set(curOption.Id); //??? the code is getting the right option, but it's not changing anything in the model
        }

        #endregion

        #region String

        internal static string GetStringBetweenCharacters(string input, string charFrom, string charTo)
        {
            //string cleanInput = CleanSheetNumber(input);

            int posFrom = input.IndexOf(charFrom);
            if (posFrom != -1) //if found char
            {
                int posTo = input.IndexOf(charTo, posFrom + 1);
                if (posTo != -1) //if found char
                {
                    return input.Substring(posFrom + 1, posTo - posFrom - 1);
                }
            }

            return string.Empty;
        }

        #endregion       

        #region Views

        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector m_colviews = new FilteredElementCollector(curDoc);
            m_colviews.OfCategory(BuiltInCategory.OST_Views);

            List<View> m_views = new List<View>();
            foreach (View x in m_colviews.ToElements())
            {
                m_views.Add(x);
            }

            return m_views;
        }

        internal static List<ViewPlan> GetAllAreaPlans(Document curDoc)
        {
            List<ViewPlan> returnList = new List<ViewPlan>();
            List<ViewPlan> viewList = GetAllViewPlans(curDoc);

            foreach (View x in viewList)
            {
                if (x.ViewType == ViewType.AreaPlan)
                {
                    returnList.Add((ViewPlan)x);
                }
            }

            return returnList;
        }

        public static List<ViewPlan> GetAllViewPlans(Document curDoc)
        {
            List<ViewPlan> returnList = new List<ViewPlan>();

            FilteredElementCollector viewCollector = new FilteredElementCollector(curDoc);
            viewCollector.OfCategory(BuiltInCategory.OST_Views);
            viewCollector.OfClass(typeof(ViewPlan)).ToElements();

            foreach (ViewPlan vp in viewCollector)
            {
                if (vp.IsTemplate == false)
                    returnList.Add(vp);
            }

            return returnList;
        }

        internal static ViewPlan GetAreaPlanByViewFamilyName(Document doc, string vftName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewPlan));

            foreach (ViewPlan curViewPlan in collector)
            {
                if (curViewPlan.ViewType == ViewType.AreaPlan)
                {
                    ViewFamilyType curVFT = doc.GetElement(curViewPlan.GetTypeId()) as ViewFamilyType;

                    if (curVFT.Name == vftName)
                        return curViewPlan;
                }
            }

            return null;
        }

        #endregion        

        #region View Templates

        public static View GetViewTemplateByName(Document curDoc, string viewTemplateName)
        {
            List<View> viewTemplateList = GetAllViewTemplates(curDoc);

            foreach (View v in viewTemplateList)
            {
                if (v.Name == viewTemplateName)
                {
                    return v;
                }
            }

            return null;
        }

        public static ViewSchedule GetViewScheduleTemplateByName(Document curDoc, string viewSchedTemplateName)
        {
            List<ViewSchedule> viewSchedTemplateList = GetAllViewScheduleTemplates(curDoc);

            foreach (ViewSchedule v in viewSchedTemplateList)
            {
                if (v.Name == viewSchedTemplateName)
                {
                    return v;
                }
            }

            return null;
        }

        private static List<ViewSchedule> GetAllViewScheduleTemplates(Document curDoc)
        {
            List<ViewSchedule> returnList = new List<ViewSchedule>();
            List<ViewSchedule> viewList = GetAllSchedules(curDoc);

            //loop through views and check if is view template
            foreach (ViewSchedule v in viewList)
            {
                if (v.IsTemplate == true)
                {
                    //add view template to list
                    returnList.Add(v);
                }
            }

            return returnList;
        }

        public static List<View> GetAllViewTemplates(Document curDoc)
        {
            List<View> returnList = new List<View>();
            List<View> viewList = GetAllViews(curDoc);

            //loop through views and check if is view template
            foreach (View v in viewList)
            {
                if (v.IsTemplate == true)
                {
                    //add view template to list
                    returnList.Add(v);
                }
            }

            return returnList;
        }

        internal static void DuplicateAndRenameSheetIndex(Document curDoc, string newFilter)
        {
            // duplicate the first schedule found with "Sheet Index" in the name
            List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Sheet Index");
            ViewSchedule dupSched = listSched.FirstOrDefault();

            if (dupSched == null)
            {
                // call another method to create one

                return; // no schedule to duplicate
            }

            ViewSchedule indexSched = curDoc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;

            // rename the duplicated schedule to the new elevation
            string originalName = indexSched.Name;
            string[] schedTitle = originalName.Split('-');

            string curTitle = schedTitle[0];

            indexSched.Name = schedTitle[0] + "- Elevation " + GlobalVars.ElevDesignation;

            Utils.SetParameterValue(indexSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

            // update the filter value to the new elevation code filter
            ScheduleFilter codeFilter = indexSched.Definition.GetFilter(0);

            if (codeFilter.IsStringValue)
            {
                codeFilter.SetValue(newFilter);
                indexSched.Definition.SetFilter(0, codeFilter);
            }
        }

        internal static void DuplicateAndConfigureVeneerSchedule(Document curDoc)
        {
            // duplicate the first schedule with "Exterior Venner Calculations" in the name
            List<ViewSchedule> listSched = Utils.GetAllScheduleByNameContains(curDoc, "Exterior Veneer Calculations");
            ViewSchedule dupSched = listSched.FirstOrDefault();

            if (dupSched == null)
            {
                // call another method to create one

                return; // no schedule to duplicate
            }

            // duplicate the schedule
            ViewSchedule veneerSched = curDoc.GetElement(dupSched.Duplicate(ViewDuplicateOption.Duplicate)) as ViewSchedule;

            // rename the duplicated schedule to the new elevation
            string originalName = veneerSched.Name;
            string[] schedTitle = originalName.Split('-');

            veneerSched.Name = schedTitle[0] + "- Elevation " + GlobalVars.ElevDesignation;

            Utils.SetParameterValue(veneerSched, "Elevation Designation", "Elevation " + GlobalVars.ElevDesignation);

            //// set the design option to the specified elevation designation
            //DesignOption curOption = Utils.getDesignOptionByName(curDoc, "Elevation : " + Globals.ElevDesignation);

            //Parameter doParam = veneerSched.get_Parameter(BuiltInParameter.VIEWER_OPTION_VISIBILITY);

            //doParam.Set(curOption.Id); //??? the code is getting the right option, but it's not changing anything in the model
        }

        internal static void CreateFloorAreaWithTag(Document curDoc, ViewPlan areaPlan, ref UV insPoint, ref XYZ tagInsert, clsAreaData areaInfo)
        {
            Area curArea = curDoc.Create.NewArea(areaPlan, insPoint);
            curArea.Number = areaInfo.Number;
            curArea.Name = areaInfo.Name;
            curArea.LookupParameter("Area Category").Set(areaInfo.Category);
            curArea.LookupParameter("Comments").Set(areaInfo.Comments);

            AreaTag tag = curDoc.Create.NewAreaTag(areaPlan, curArea, insPoint);
            tag.TagHeadPosition = tagInsert;
            tag.HasLeader = false;

            UV offset = new UV(0, 8);
            insPoint = insPoint.Subtract(offset);

            XYZ tagOffset = new XYZ(0, 8, 0);
            tagInsert = tagInsert.Subtract(tagOffset);

            if (areaInfo.Ratio != 99)
            {
                curArea.LookupParameter("150 Ratio").Set(areaInfo.Ratio);
            }
        }

        #endregion

        internal static ColorFillScheme GetColorFillSchemeByName(Document curDoc, string schemeName, AreaScheme areaScheme)
        {
            try
            {
                ColorFillScheme colorfill = new FilteredElementCollector(curDoc)
               .OfCategory(BuiltInCategory.OST_ColorFillSchema)
               .Cast<ColorFillScheme>()
               .Where(x => x.Name.Equals(schemeName) && x.AreaSchemeId.Equals(areaScheme.Id))
               .First();

                return colorfill;
            }
            catch
            {
                return null;
            }
        }

        internal static void AddColorLegend(View view, ColorFillScheme scheme)
        {
            ElementId areaCatId = new ElementId(BuiltInCategory.OST_Areas);
            ElementId curLegendId = view.GetColorFillSchemeId(areaCatId);

            if (curLegendId == ElementId.InvalidElementId)
                view.SetColorFillSchemeId(areaCatId, scheme.Id);

            ColorFillLegend.Create(view.Document, view.Id, areaCatId, XYZ.Zero);
        }
    }
}

