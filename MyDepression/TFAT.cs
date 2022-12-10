using Excel = Microsoft.Office.Interop.Excel;

using System.Collections.Generic;
using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using System.IO;
using Microsoft.Office.Tools.Excel;
using ListObject = Microsoft.Office.Interop.Excel.ListObject;
using System.Windows.Forms;
using System.Runtime.CompilerServices;
using static System.Net.Mime.MediaTypeNames;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MyDepression
{
    public static class TFAT
    {
        public enum Column { First, Last, Email, Rank, Training, Date };
        public static string[] Ranks { get => new string[] { "AB", "A1C", "SrA", "SSgt", "TSgt", "MSgt", "SMSgt", "CMSgt", "1st Lt", "2d Lt", "Capt", "Maj", "Lt Col", "Col" }; }

        public readonly struct Person
        {
            public string FirstName { get; }
            public string LastName { get; }
            public string Email { get; }
            public string Rank { get; }
            public string Organization { get; }
            public bool IsMil { get; }

            public Person(string firstName, string lastName, string email, string rank, string organization)
            {
                FirstName = firstName ?? throw new ArgumentNullException(nameof(firstName));
                LastName = lastName ?? throw new ArgumentNullException(nameof(lastName));
                Email = email ?? throw new ArgumentNullException(nameof(email));
                Rank = rank ?? throw new ArgumentNullException(nameof(rank));
                Organization = organization ?? throw new ArgumentNullException(nameof(organization));

                IsMil = false;

                foreach (var r in Ranks)
                    IsMil |= r.ToLower() == Rank.ToLower();
            }
        }
        public readonly struct TFATRecord
        {
            public Person Member { get; }
            public string Training { get; }
            public DateTime? Date { get; }

            public TFATRecord(Person member, string training, DateTime? date)
            {
                Member = member;
                Training = training ?? throw new ArgumentNullException(nameof(training));
                Date = date;
            }
        }
        [Serializable]
        public class TFATTraining
        {
            public string SafeName { get; set; }
            public string Name { get; set; }
            public int Suspense { get; set; }
            public bool ReqMil { get; set; }
            public bool ReqCiv { get; set; }

            public TFATTraining() { }
            public TFATTraining(string safeName, string name, int suspense, bool reqMil, bool reqCiv)
            {
                SafeName = safeName ?? throw new ArgumentNullException(nameof(safeName));
                Name = name ?? throw new ArgumentNullException(nameof(name));
                Suspense = suspense;
                ReqMil = reqMil;
                ReqCiv = reqCiv;
            }

            public override string ToString() => SafeName;
        }

        public static List<TFATTraining> LoadTrainings(string path)
        {
            List<TFATTraining> output = null;

            XmlSerializer xml = new XmlSerializer(typeof(List<TFATTraining>));

            using (FileStream fs = File.OpenRead(path))
                output = xml.Deserialize(fs) as List<TFATTraining>;

            return output;
        }
        public static void SaveTrainings(string path, List<TFATTraining> trainings)
        {
            XmlSerializer xml = new XmlSerializer(trainings.GetType());

            using (FileStream fs = File.Create(path))
                xml.Serialize(fs, trainings);
        }

        public static List<TFATRecord> ParseTFAT(Excel.Worksheet sheet, Dictionary<Column, string> colNames)
        {
            // Warning, excel docs are 1 based

            var output = new List<TFATRecord>();

            Range data = sheet.UsedRange;

            int rowCnt = data.Rows.Count;
            int colCnt = data.Columns.Count;

            // Find col headers
            int first = -1,
                last = -1,
                email = -1,
                rank = -1,
                trg = -1, 
                date = -1;

            for (int i = 1; i < colCnt + 1; i++)
            {
                string header = (data.Cells[1, i] as Range).Text;

                if (colNames[Column.Email] == header)
                    email = i;
                else if (colNames[Column.Date] == header)
                    date = i;
                else if (colNames[Column.Training] == header)
                    trg = i;
                else if (colNames[Column.First] == header)
                    first = i;
                else if (colNames[Column.Last] == header)
                    last = i;
                else if (colNames[Column.Rank] == header)
                    rank = i;
            }

            if (email == -1)
                throw new Exception("Failed to locate Email column. Please update your association settings!");
            if (trg == -1)
                throw new Exception("Failed to locate Training Name column. Please update your association settings!");
            if (date == -1)
                throw new Exception("Failed to locate Completion Date column. Please update your association settings!");
            if (first == -1)
                throw new Exception("Failed to locate First Name column. Please update your association settings!");
            if (last == -1)
                throw new Exception("Failed to locate Last Name column. Please update your association settings!");
            if (rank == -1)
                throw new Exception("Failed to locate Rank column. Please update your association settings!");

            string getText(int row, int col)
            {
                try
                {
                    return (data.Cells[row, col] as Range).Text;
                } catch
                {
                    return "";
                }
            }

            // Parse records
            for (int i = 2; i < rowCnt + 1; i++)
            {
                DateTime? dateTime = null;
                DateTime parse;

                if (DateTime.TryParse(getText(i, date), out parse))
                    dateTime = parse;

                Person member = new Person(getText(i, first), getText(i, last), getText(i, email), getText(i, rank), "");
                output.Add(new TFATRecord(member, getText(i, trg), dateTime));
            }

            return output;
        }
    
        public static void UpdateTFATTable(ListObject table, List<TFATRecord> records, List<TFATTraining> trainings)
        {
            // Validate table
            var cols = new Dictionary<string, int>();

            int idx = 1;
            foreach (var header in table.HeaderRowRange.Cells)
            {
                try
                {
                    string text = (header as Range).Text;
                    if (text == "Email")
                    {
                        cols.Add("Email", idx);
                    }
                    else if (text == "Affiliation")
                    {
                        cols.Add("Affiliation", idx);
                    }
                    else
                    {
                        foreach (var trg in trainings)
                        {
                            if (text == trg.SafeName)
                            {
                                cols.Add(trg.SafeName, idx);
                                break;
                            }
                        }
                    }
                }
                finally {
                    idx++;
                }
            }

            if (!cols.ContainsKey("Email"))
            {
                MessageBox.Show("Failed to find a column for \"Email\" in the tracker table. You must add one to continue.");
                return;
            }
            if (!cols.ContainsKey("Affiliation"))
            {
                MessageBox.Show("Failed to find a column for \"Affiliation\" in the tracker table. You must add one to continue.");
                return;
            }

            foreach (var trg in trainings)
            {
                if (!cols.ContainsKey(trg.SafeName))
                {
                    MessageBox.Show("Failed to find a column for \"" + trg.SafeName + "\"  in the tracker table. You must add one or remove it from the required trainings list.");
                    return;
                }
            }
            
            // People list
            var pplData = new Dictionary<string, Range>();

            foreach (var row in table.ListRows)
            {
                try
                {
                    Range range = (row as ListRow).Range;
                    
                    pplData.Add((range.Cells[1, cols["Email"]] as Range).Text, range);
                } catch (Exception ex)
                {
                    MessageBox.Show("Error:\n" + ex.Message);
                }
            }

            // Name to safename map
            var safeNames = new Dictionary<string, string>();
            foreach (var trg in trainings)
                safeNames.Add(trg.Name, trg.SafeName);

            // Trg name to trg map
            var trgMap = new Dictionary<string, TFATTraining>();
            foreach (var trg in trainings)
                trgMap.Add(trg.Name, trg);

            // Update trainings
            foreach (var rec in records)
            {
                try
                {
                    // Skip trainings that are not in the required list
                    if (!safeNames.ContainsKey(rec.Training))
                        continue;

                    Range cell = pplData[rec.Member.Email].Cells[1, cols[safeNames[rec.Training]]];

                    DateTime? date = rec.Date;

                    DateTime val;
                    if (DateTime.TryParse(cell.Text, out val))
                    {
                        date = val;

                        if (rec.Date.HasValue && rec.Date > date)
                            date = rec.Date;
                    }

                    if (date.HasValue)
                    {
                        var trg = trgMap[rec.Training];

                        cell.Value2 = date.Value.ToShortDateString();

                        DateTime sus = date.Value.AddMonths(trg.Suspense);

                        int diff = (DateTime.Now - sus).Days;


                        if (diff >= 0)
                            cell.Style = "Bad";
                        else if (diff >= -30)
                            cell.Style = "Neutral";
                        else
                            cell.Style = "Good";
                    }

                } catch (Exception ex)
                {
                    MessageBox.Show("Error:\n" + ex.Message);
                }
            }
            

            // Check non updated trainings
            foreach (var pers in pplData.Values)
            {
                string affil = ((pers.Cells[1, cols["Affiliation"]] as Range).Text as string).ToLower();
                bool mil = affil == "mil";

                if (affil != "civ" && affil != "mil")
                {
                    MessageBox.Show("Affiliations must be either \"Mil\" or \"Civ\".");
                    continue;
                }

                foreach (var trg in trgMap.Values)
                {
                    try
                    {
                        Range cell = pers.Cells[1, cols[trg.SafeName]] as Range;

                        if ((mil && !trg.ReqMil) || (!mil && !trg.ReqCiv))
                        {
                            cell.Style = "Good";
                            cell.Value2 = "N/A";
                        }
                        else if (cell.Text as string == "")
                        {
                            cell.Style = "Bad";
                            cell.Value2 = "Not Complete";
                        }
                    } catch { }
                }
            }
            /*
            Outlook.Application oApp = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "This is the subject";
            mailItem.To = "someone@example.com";
            mailItem.Body = "This is the message.";
            mailItem.Display(false);
            */
        }
    }
}