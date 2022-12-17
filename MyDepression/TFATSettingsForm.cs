using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;
using static MyDepression.TFAT;

namespace MyDepression
{
    public partial class TFATSettingsForm : Form
    {
        [Serializable]
        public class TFATSettings
        {
            [Serializable]
            public struct ColPair
            {
                public Column Col;
                public string Val;

                public ColPair(Column col, string val)
                {
                    Col = col;
                    Val = val ?? throw new ArgumentNullException(nameof(val));
                }
            }

            public List<TFATTraining> Trainings;
            [XmlIgnore]
            public Dictionary<Column, string> Associations
            {
                get
                {
                    var output = new Dictionary<Column, string>();

                    foreach (var item in AssociationsList)
                        output.Add(item.Col, item.Val);

                    return output;
                }
            }
            public List<ColPair> AssociationsList;
            public string EmailSubj;
            public string EmailBody;

            public TFATSettings() { }
            public TFATSettings(List<TFATTraining> trainings, List<KeyValuePair<Column, string>> associations, string emailSubj, string emailBody)
            {
                Trainings = trainings ?? throw new ArgumentNullException(nameof(trainings));
                EmailSubj = emailSubj ?? throw new ArgumentNullException(nameof(emailSubj));
                EmailBody = emailBody ?? throw new ArgumentNullException(nameof(emailBody));

                AssociationsList = new List<ColPair>();
                foreach (var val in associations)
                    AssociationsList.Add(new ColPair(val.Key, val.Value));
            }

            public static TFATSettings Load(string path)
            {
                TFATSettings output = null;

                XmlSerializer xml = new XmlSerializer(typeof(TFATSettings));

                using (FileStream fs = File.OpenRead(path))
                    output = xml.Deserialize(fs) as TFATSettings;

                return output;
            }
            public void Save(string path)
            {
                XmlSerializer xml = new XmlSerializer(GetType());

                using (FileStream fs = File.Create(path))
                    xml.Serialize(fs, this);
            }
        }

        BindingList<TFATTraining> trainings = new BindingList<TFATTraining>();
        public List<TFATTraining> Trainings() => trainings.ToList();
        public Dictionary<Column, string> Associations() => new Dictionary<TFAT.Column, string>()
        {
            { TFAT.Column.Email, emailBox.Text },
            { TFAT.Column.Training, nameBox.Text },
            { TFAT.Column.Date, dateBox.Text },
            { TFAT.Column.First, fnBox.Text },
            { TFAT.Column.Last, lnBox.Text },
            { TFAT.Column.Rank, rankBox.Text },
        };
        public string EmailSubject() => subjBox.Text;
        public string EmailBody() => bodyBox.Text;
        void loadSettings(string path)
        {
            try
            {
                TFATSettings settings = TFATSettings.Load(path);
                trainings = new BindingList<TFATTraining>(settings.Trainings);

                foreach (var pair in settings.AssociationsList)
                {
                    switch (pair.Col)
                    {
                        case TFAT.Column.Email:
                            emailBox.Text = pair.Val;
                            break;
                        case TFAT.Column.Training:
                            nameBox.Text = pair.Val;
                            break;
                        case TFAT.Column.Date:
                            dateBox.Text = pair.Val;
                            break;
                        case TFAT.Column.First:
                            fnBox.Text = pair.Val;
                            break;
                        case TFAT.Column.Last:
                            lnBox.Text = pair.Val;
                            break;
                    }
                }

                subjBox.Text = settings.EmailSubj;
                bodyBox.Text = Encoding.Unicode.GetString(Convert.FromBase64String(settings.EmailBody));
            }
            catch
            {

            }

            if (trainings.Count == 0)
                trgAddBtn_Click(null, null);

            trgList.DataSource = trainings;
        }
        public TFATSettingsForm()
        {
            InitializeComponent();

            loadSettings("TFAT_Settings.xml");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void TFATAssociation_Load(object sender, EventArgs e)
        {

        }

        private void TFATAssociation_FormClosing(object sender, FormClosingEventArgs e)
        {
            new TFATSettings(Trainings(), Associations().ToList(), subjBox.Text, Convert.ToBase64String(Encoding.Unicode.GetBytes(bodyBox.Text))).Save("TFAT_Settings.xml");
        }

        private void trgAddBtn_Click(object sender, EventArgs e)
        {
            trainings.Add(new TFAT.TFATTraining("Short Name", "Full Name", 12, true, true));
        }

        private void trgList_SelectedIndexChanged(object sender, EventArgs e)
        {
            trgViewPanel.Visible = trgList.SelectedItem != null;

            trgSName.DataBindings.Clear();
            trgFName.DataBindings.Clear();
            susBox.DataBindings.Clear();
            cbReqCiv.DataBindings.Clear();
            cbReqMil.DataBindings.Clear();

            if (trgList.SelectedItem != null)
            {
                trgSName.DataBindings.Add(new Binding("Text", trgList.SelectedItem, "SafeName", false, DataSourceUpdateMode.OnPropertyChanged));
                trgFName.DataBindings.Add(new Binding("Text", trgList.SelectedItem, "Name", false, DataSourceUpdateMode.OnPropertyChanged));
                susBox.DataBindings.Add(new Binding("Value", trgList.SelectedItem, "Suspense", false, DataSourceUpdateMode.OnPropertyChanged));
                cbReqCiv.DataBindings.Add(new Binding("Checked", trgList.SelectedItem, "ReqCiv", false, DataSourceUpdateMode.OnPropertyChanged));
                cbReqMil.DataBindings.Add(new Binding("Checked", trgList.SelectedItem, "ReqMil", false, DataSourceUpdateMode.OnPropertyChanged));
            }
        }

        private void trgSName_TextChanged(object sender, EventArgs e)
        {
            trgList.Refresh();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            trainings.Remove((TFAT.TFATTraining)trgList.SelectedItem);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                var sfd = new OpenFileDialog() { Filter = "XML Document|*.xml" };

                if (sfd.ShowDialog() == DialogResult.OK)
                    loadSettings(sfd.FileName);

                trgList.DataSource = null;
                trgList.DataSource = trainings;

            }
            catch (Exception ex)
            {
                new ExceptionForm(ex).ShowDialog();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                var sfd = new SaveFileDialog() { Filter = "XML Document|*.xml" };

                if (sfd.ShowDialog() == DialogResult.OK)
                    new TFATSettings(Trainings(), Associations().ToList(), subjBox.Text, Convert.ToBase64String(Encoding.Unicode.GetBytes(bodyBox.Text))).Save(sfd.FileName);

            }
            catch (Exception ex)
            {
                new ExceptionForm(ex).ShowDialog();
            }
        }
    }
}