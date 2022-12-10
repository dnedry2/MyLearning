using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace MyDepression
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
            /*
            List<TFAT.TFATTraining> trainings = new List<TFAT.TFATTraining>
            {
                new TFAT.TFATTraining("Force Protection", "Force Protection (ZZ133079)", 12, true, false),
                new TFAT.TFATTraining("Cyber Awareness", "Cyber Awareness Challenge 2021 (ZZ133098)", 12, true, true),
                new TFAT.TFATTraining("Cyber Awareness", "Cyber Awareness Challenge 2022 (ZZ133098)", 12, true, true)
            };

            TFAT.SaveTrainings("trainings.xml", trainings);
            */
        }

        List<TFAT.TFATRecord> LoadRecords()
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel|*.xls;*.xlsx;*.csv" };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var transcript = Globals.ThisAddIn.Application.Workbooks.Open(ofd.FileName);

                    var records = TFAT.ParseTFAT(Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet, TFATSettings.Associations());

                    transcript.Close(false);
                    return records;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to load document!\n" + ex.Message);
                }
            }

            return null;
        }

        List<TFAT.TFATTraining> LoadTrainings()
        {
            List<TFAT.TFATTraining> trainings;
            try
            {
                trainings = TFAT.LoadTrainings("trainings.xml");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load trainings!\n" + ex.Message);
                return null;
            }

            if (trainings.Count == 0)
            {
                MessageBox.Show("No trainings. You need to setup required trainings first.");
                return null;
            }

            return trainings;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MyDepression.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        TFATSettingsForm TFATSettings = new TFATSettingsForm();

        public void button1_Click(Office.IRibbonControl control)
        {
            Excel.Worksheet tracker = null;

            foreach (var sheet in Globals.ThisAddIn.Application.Worksheets)
            {
                Excel.Worksheet ws = sheet as Excel.Worksheet;

                if (ws != null && ws.Name == "TFAT Tracker")
                    tracker = ws;
            }

            if (tracker == null)
            {
                MessageBox.Show("TFAT Tracker worksheet was not found. Create one before updating.");
                return;
            }

            ListObject table;

            try
            {
                table = tracker.ListObjects["TFAT Table"];
            }
            catch
            {
                MessageBox.Show("TFAT Table was not found.");
                return;
            }


            var trainings = LoadTrainings();
            var records = LoadRecords();

            if (records == null || trainings == null)
                return;

            TFAT.UpdateTFATTable(table, records, trainings);
        }
        public void button2_Click(Office.IRibbonControl control)
        {
            Excel.Worksheet tracker = null;

            foreach (var sheet in Globals.ThisAddIn.Application.Worksheets)
            {
                Excel.Worksheet ws = sheet as Excel.Worksheet;

                if (ws != null && ws.Name == "TFAT Tracker")
                    tracker = ws;
            }

            if (tracker == null)
            {
                MessageBox.Show("TFAT Tracker worksheet was not found. Create one before updating.");
                return;
            }

            ListObject table;

            try
            {
                table = tracker.ListObjects["TFAT Table"];
            }
            catch
            {
                MessageBox.Show("TFAT Table was not found.");
                return;
            }

            var trainings = LoadTrainings();

            if (trainings == null)
                return;

            TFAT.EmailNotification(table, trainings, TFATSettings.EmailSubject(), TFATSettings.EmailBody());
        }

        public void button3_Click(Office.IRibbonControl control)
        {
            TFATSettings.ShowDialog();
        }

        public void button4_Click(Office.IRibbonControl control)
        {
            var trainings = LoadTrainings();
            var records = LoadRecords();

            if (records == null || trainings == null)
                return;


            Excel.Worksheet tracker = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();

            try
            {
                tracker.Name = "TFAT Tracker";
            }
            catch
            {
                MessageBox.Show("Failed to create TFAT Tracker worksheet. (Does it already exist?)");
                return;
            }

            tracker.Activate();

            // Setup headers
            tracker.Cells[1, 1] = "Last Name";
            tracker.Cells[1, 2] = "First Name";
            tracker.Cells[1, 3] = "Rank";
            tracker.Cells[1, 4] = "Affiliation";
            tracker.Cells[1, 5] = "Organization";
            tracker.Cells[1, 6] = "Email";

            // Find unique trainings
            var trgs = new HashSet<string>();
            foreach (var trg in trainings)
                trgs.Add(trg.SafeName);

            int i = 0;
            foreach (string trg in trgs)
                tracker.Cells[1, 7 + i++] = trg;


            // Add people

            // Find unique people
            var people = new HashSet<TFAT.Person>();
            foreach (var rec in records)
                people.Add(rec.Member);

            // Add people to tracker
            i = 0;
            foreach (var person in people)
            {
                tracker.Cells[2 + i, 1] = person.LastName;
                tracker.Cells[2 + i, 2] = person.FirstName;
                tracker.Cells[2 + i, 3] = person.Rank;
                tracker.Cells[2 + i, 4] = person.IsMil ? "Mil" : "Civ";
                tracker.Cells[2 + i, 5] = person.Organization;
                tracker.Cells[2 + i, 6] = person.Email;

                i++;
            }

            tracker.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, tracker.UsedRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name = "TFAT Table";
            tracker.ListObjects["TFAT Table"].TableStyle = "TableStyleMedium21";

            TFAT.UpdateTFATTable(tracker.ListObjects["TFAT Table"], records, trainings);
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}