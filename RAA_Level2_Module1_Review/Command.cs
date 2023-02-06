#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace RAA_Level2_Module1_Review
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // open form
            MyForm currentForm = new MyForm()
            {
                Width = 500,
                Height = 365,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            if(currentForm.ShowDialog() == false)
            {
                return Result.Cancelled;
            }

            // 4. get form data and do something
            string filePath = currentForm.GetFilePath(); 

            if(filePath == null)
            {
                return Result.Cancelled;
            }

            // read CSV file
            string[] levelArray = System.IO.File.ReadAllLines(filePath);

            // read CSV data into a list
            List<string[]> levelData = new List<string[]>();

            foreach(string levelString in levelArray)
            {
                string[] cellData = levelString.Split(',');
                levelData.Add(cellData);
            }

            // remove header
            levelData.RemoveAt(0);

            Transaction t = new Transaction(doc);
            t.Start("Project setup");

            // create levels and views
            foreach (string[] currentLevelData in levelData)
            {
                // set variables
                string levelName = currentLevelData[0];
                string elevImperialString = currentLevelData[1];
                string elevMetricString = currentLevelData[2];

                double levelElevation = 0;

                if(currentForm.GetUnits() == "imperial")
                {
                    // imperial
                    double.TryParse(elevImperialString, out levelElevation);
                }
                else
                {
                    // metric
                    double metricConvert = 0;
                    double.TryParse(elevMetricString, out metricConvert);
                    levelElevation = metricConvert * 3.28084;
                }

                // create level
                Level currentLevel = Level.Create(doc, levelElevation);
                currentLevel.Name = levelName;

                // create views
                if(currentForm.CreateFloorPlan() == true)
                {
                    ViewFamilyType planVFT = GetViewFamilyTypeByName(doc, "Floor Plan", ViewFamily.FloorPlan);

                    ViewPlan currentPlan = ViewPlan.Create(doc, planVFT.Id, currentLevel.Id);
                }

                if(currentForm.CreateCeilingPlan() == true)
                {
                    ViewFamilyType ceilingPlanVFT = GetViewFamilyTypeByName(doc, "Ceiling Plan", ViewFamily.CeilingPlan);

                    ViewPlan currentCeilingPlan = ViewPlan.Create(doc, ceilingPlanVFT.Id, currentLevel.Id);
                }
            }

            t.Commit();
            t.Dispose();

            TaskDialog.Show("Complete", "Created " + levelData.Count + " levels and views.");

            return Result.Succeeded;
        }

        private ViewFamilyType GetViewFamilyTypeByName(Document doc, string vftName, ViewFamily viewFamily)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));

            foreach(ViewFamilyType currentVFT in collector)
            {
                if (currentVFT.Name == vftName && currentVFT.ViewFamily == viewFamily)
                    return currentVFT;
            }

            return null;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
