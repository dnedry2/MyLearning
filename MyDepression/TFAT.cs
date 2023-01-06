using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using Chart = Microsoft.Office.Tools.Excel.Chart;
using Excel = Microsoft.Office.Interop.Excel;
using Exception = System.Exception;
using ListObject = Microsoft.Office.Interop.Excel.ListObject;
using Outlook = Microsoft.Office.Interop.Outlook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

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
            public string Code { get; set; } = "";

            public TFATTraining() { }
            public TFATTraining(string safeName, string name, int suspense, bool reqMil, bool reqCiv, string code)
            {
                SafeName = safeName ?? throw new ArgumentNullException(nameof(safeName));
                Name = name ?? throw new ArgumentNullException(nameof(name));
                Suspense = suspense;
                ReqMil = reqMil;
                ReqCiv = reqCiv;
                Code = code ?? throw new ArgumentNullException(nameof(name));
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

        static Dictionary<string, int> validate(ListObject table, List<TFATTraining> trainings)
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
                    else if (text == "Rank")
                    {
                        cols.Add("Rank", idx);
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
                finally
                {
                    idx++;
                }
            }

            if (!cols.ContainsKey("Email"))
            {
                MessageBox.Show("Failed to find a column for \"Email\" in the tracker table. You must add one to continue.");
                return null;
            }
            if (!cols.ContainsKey("Affiliation"))
            {
                MessageBox.Show("Failed to find a column for \"Affiliation\" in the tracker table. You must add one to continue.");
                return null;
            }

            foreach (var trg in trainings)
            {
                if (!cols.ContainsKey(trg.SafeName))
                {
                    MessageBox.Show("Failed to find a column for \"" + trg.SafeName + "\"  in the tracker table. You must add one or remove it from the required trainings list.");
                    return null;
                }
            }

            return cols;
        }

        static Dictionary<string, Range> getPeopleRows(ListObject table, Dictionary<string, int> cols)
        {
            var pplData = new Dictionary<string, Range>();

            foreach (var row in table.ListRows)
            {
                try
                {
                    Range range = (row as ListRow).Range;

                    pplData.Add((range.Cells[1, cols["Email"]] as Range).Text, range);
                }
                catch (Exception ex)
                {
                    new ExceptionForm(ex).ShowDialog();
                }
            }

            return pplData;
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
                }
                catch
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
            var cols = validate(table, trainings);

            if (cols == null)
                return;

            // People list
            var pplData = getPeopleRows(table, cols);

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

                    // Skip people that are not in the people list
                    if (!pplData.ContainsKey(rec.Member.Email))
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
                        cell.Value2 = date.Value.ToShortDateString();

                }
                catch (Exception ex)
                {
                    new ExceptionForm(ex).ShowDialog();
                }
            }

            // Check non updated trainings
            foreach (var pers in pplData.Values)
            {
                // Skip blank lines
                if (pers.Text as string == "")
                    continue;

                string affil = ((pers.Cells[1, cols["Affiliation"]] as Range).Text as string).ToLower();
                bool mil = affil == "mil";

                if (affil != "civ" && affil != "mil")
                {
                    MessageBox.Show("Affiliations must be either \"Mil\" or \"Civ\". Was \"" + affil + "\"");
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
                        } else
                        {
                            DateTime? date = null;

                            DateTime val;
                            if (DateTime.TryParse(cell.Value.ToString(), out val))
                                date = val;

                            if (date.HasValue)
                            {
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
                        }
                        /*
                        if (cell.Text != null && cell.Text == "Not Complete") {
                            cell.Value2 = $"=HYPERLINK(\"https://lms-jets.cce.af.mil/moodle/user/view.php?course={trg.Code}&email={(pers.Cells[1, cols["Email"]].Text as string).Replace("@", "%40")}\",\"Not Complete\")";
                        }
                        */
                    }
                    catch { }
                }
            }
        }

        static string parseEmail(string message, TFATTraining trg)
        {
            string dateStr = "Morning";

            if (DateTime.Now.Hour >= 11)
                dateStr = "Afternoon";
            if (DateTime.Now.Hour >= 18)
                dateStr = "Evening";

            return message.Replace("$ShortName", trg.SafeName)
                          .Replace("$FullName", trg.Name)
                          .Replace("$Suspense", Convert.ToString(trg.Suspense))
                          .Replace("$Time", dateStr)
                          .Replace("$URL", $"https://lms-jets.cce.af.mil/moodle/user/view.php?course={trg.Code}");
        }

        public static void EmailNotification(ListObject table, List<TFATTraining> trainings, string subj, string body)
        {
            // Validate table
            var cols = validate(table, trainings);

            if (cols == null)
                return;

            Outlook.Application oApp = new Outlook.Application();

            Dictionary<string, TFATTraining> uniqueTrgs = new Dictionary<string, TFATTraining>();
            foreach (var trg in trainings)
                uniqueTrgs[trg.SafeName] = trg;

            // People list
            var pplData = getPeopleRows(table, cols);

            // Generate emails
            foreach (var trg in uniqueTrgs.Values)
            {
                string to = "";

                foreach (var p in pplData)
                {
                    string style = (p.Value.Cells[1, cols[trg.SafeName]].Style as Style).Value;
                    
                    if (style == "Bad")
                        to += p.Key + ";";
                }

                if (to == "")
                    continue;


                Outlook.MailItem mailItem = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = parseEmail(subj, trg);
                mailItem.To = to;
                mailItem.Body = parseEmail(body, trg);
                mailItem.Display(false);
            }
        }

        static int findRow(int points, Worksheet sheet)
        {
            int idx = 1;

            while (true)
            {
                if ((sheet.Rows[idx] as Range).Top > points)
                    break;

                idx++;
            }


            return idx;
        }

        static string getExcelCol(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        public static void CalculateStatistics(ListObject table, Worksheet output, List<TFATTraining> trainings)
        {
            var cols = validate(table, trainings);

            if (cols == null)
                return;

            var pplData = getPeopleRows(table, cols);

            var uniqueTrgs = new Dictionary<string, TFATTraining>();
            foreach (var trg in trainings)
                uniqueTrgs[trg.SafeName] = trg;

            var trgStats = new List<(TFATTraining trg, int req, int bad, int warn, int good)>();
            foreach (var trg in uniqueTrgs.Values)
                trgStats.Add((trg, 0, 0, 0, 0));

            (int t, int g) milStats = (0, 0),
                           civStats = (0, 0);

            var rankTotals = new Dictionary<string, (int t, int g)>();

            for (int i = 0; i < trgStats.Count; i++)
            {
                var trg = trgStats[i];

                foreach (var p in pplData)
                {
                    var cell = p.Value.Cells[1, cols[trg.trg.SafeName]] as Range;
                    string style = (cell.Style as Style).Value;
                    string text = cell.Text;

                    if (text == "N/A")
                        continue;

                    trg.req++;

                    bool mil = (p.Value.Cells[1, cols["Affiliation"]] as Range).Text == "Mil";
                    string rank = (p.Value.Cells[1, cols["Rank"]] as Range).Text;

                    if (mil)
                        milStats.t++;
                    else
                        civStats.t++;

                    if (!rankTotals.ContainsKey(rank))
                        rankTotals.Add(rank, (0, 0));

                    var cRank = rankTotals[rank];
                    cRank.t++;

                    if (style == "Bad")
                        trg.bad++;
                    if (style == "Neutral")
                        trg.warn++;
                    

                    if (style == "Good" || style == "Neutral")
                    {
                        trg.good++;

                        if (mil)
                            milStats.g++;
                        else
                            civStats.g++;

                        cRank.g++;
                    }

                    rankTotals[rank] = cRank;
                }

                trgStats[i] = trg;
            }


            // Training Stats

            // Table
            output.Cells[1, 1] = "Name";
            output.Cells[1, 2] = "Required";
            output.Cells[1, 3] = "Completed";
            output.Cells[1, 4] = "Overdue";
            output.Cells[1, 5] = "Due Soon";
            output.Cells[1, 6] = "Percent Complete";

            int idx = 2;
            foreach (var trg in trgStats)
            {
                output.Cells[idx, 1] = trg.trg.SafeName;
                output.Cells[idx, 2] = trg.req;
                output.Cells[idx, 3] = trg.good;
                output.Cells[idx, 4] = trg.bad;
                output.Cells[idx, 5] = trg.warn;
                output.Cells[idx, 6] = "=INDIRECT(\"C\" & ROW()) / INDIRECT(\"B\" & ROW())";
                (output.Cells[idx, 6] as Range).NumberFormat = "0.0%";

                idx++;
            }

            output.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, output.UsedRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name = "TFAT Stats";
            output.ListObjects["TFAT Stats"].TableStyle = "TableStyleMedium2";


            // Charts
            int ccCnt = trgStats.Count;
            int chCols = (int)Math.Ceiling(ccCnt / 2.0);

            int bottom = 0;

            var charts = output.ChartObjects(Type.Missing) as ChartObjects;

            {
                var chart = charts.Add(0, output.UsedRange.Rows.Height, 256, 256).Chart;

                chart.SetSourceData(output.Range["C2", "D2"], XlRowCol.xlRows);

                chart.ChartType = XlChartType.xlPie;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Total";

                bottom = (int)output.UsedRange.Rows.Height + 256;
            }

            for (int i = 0; i < ccCnt; i++)
            {
                int x, y;
                if (i < chCols)
                {
                    x = 256 + (i * 128);
                    y = (int)output.UsedRange.Rows.Height;
                } else
                {
                    x = 256 + ((i - chCols) * 128);
                    y = (int)output.UsedRange.Rows.Height + 128;
                }

                var chart = charts.Add(x, y, 128, 128).Chart;

                chart.SetSourceData(output.Range[$"C{i + 2}", $"D{i + 2}"], XlRowCol.xlRows);

                chart.ChartType = XlChartType.xlPie;
                chart.HasTitle = true;
                chart.ChartTitle.Text = trgStats[i].trg.SafeName;
            }


            // Rank Stats

            // Table
            int last = findRow(bottom, output);

            output.Cells[last + 1, 1] = "Status";
            output.Cells[last + 2, 1] = "Completed";
            output.Cells[last + 3, 1] = "Overdue";
            output.Cells[last + 4, 1] = "Percent Complete";


            output.Cells[last + 1, 2] = "Mil";
            output.Cells[last + 2, 2] = milStats.g;
            output.Cells[last + 3, 2] = milStats.t - milStats.g;
            output.Cells[last + 4, 2] = (double)milStats.g / milStats.t;
            (output.Cells[last + 4, 2] as Range).NumberFormat = "0.0%";

            output.Cells[last + 1, 3] = "Civ";
            output.Cells[last + 2, 3] = civStats.g;
            output.Cells[last + 3, 3] = civStats.t - milStats.g;
            output.Cells[last + 4, 3] = (double)milStats.g / milStats.t;
            (output.Cells[last + 4, 3] as Range).NumberFormat = "0.0%";

            idx = 0;
            foreach (var rnk in rankTotals.Keys) {
                output.Cells[last + 1, idx + 4] = rnk;
                output.Cells[last + 2, idx + 4] = rankTotals[rnk].g;
                output.Cells[last + 3, idx + 4] = rankTotals[rnk].t - rankTotals[rnk].g;
                output.Cells[last + 4, idx + 4] = (double)rankTotals[rnk].g / rankTotals[rnk].t;
                (output.Cells[last + 4, idx + 4] as Range).NumberFormat = "0.0%";

                idx++;
            }

            output.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, output.Range[$"A{last + 1}:{getExcelCol(idx + 3)}{last + 4}"], Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name = "TFAT Rank Stats";
            output.ListObjects["TFAT Rank Stats"].TableStyle = "TableStyleMedium2";


            // Charts
            ccCnt = rankTotals.Count + 2;
            chCols = (int)Math.Ceiling(ccCnt / 2.0);

            {
                var chart = charts.Add(0, output.UsedRange.Rows.Height, 128, 128).Chart;

                chart.SetSourceData(output.Range[$"B{last + 2}:B{last + 3}"], XlRowCol.xlColumns);

                chart.ChartType = XlChartType.xlPie;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Mil";

                bottom = (int)output.UsedRange.Rows.Height + 128;
            }
            {
                var chart = charts.Add(128, output.UsedRange.Rows.Height, 128, 128).Chart;

                chart.SetSourceData(output.Range[$"C{last + 2}:C{last + 3}"], XlRowCol.xlColumns);

                chart.ChartType = XlChartType.xlPie;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Civ";

                bottom = (int)output.UsedRange.Rows.Height + 128;
            }

            for (int i = 2; i < ccCnt; i++)
            {
                int x, y;
                if (i < chCols)
                {
                    x = (i * 128);
                    y = (int)output.UsedRange.Rows.Height;
                }
                else
                {
                    x = ((i - chCols) * 128);
                    y = (int)output.UsedRange.Rows.Height + 128;
                }

                var chart = charts.Add(x, y, 128, 128).Chart;

                chart.SetSourceData(output.Range[$"{getExcelCol(i - 2 + 4)}{last + 2}", $"{getExcelCol(i - 2 + 4)}{last + 3}"], XlRowCol.xlColumns);

                chart.ChartType = XlChartType.xlPie;
                chart.HasTitle = true;
                chart.ChartTitle.Text = rankTotals.Keys.ToList()[i - 2];
            }
        }
    }
}