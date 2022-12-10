using MyDepression.Properties;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyDepression
{
    public partial class TFATSettings : Form
    {
        BindingList<TFAT.TFATTraining> trainings = new BindingList<TFAT.TFATTraining>();
        public List<TFAT.TFATTraining> Trainings() => trainings.ToList();
        public Dictionary<TFAT.Column, string> Associations() => new Dictionary<TFAT.Column, string>()
        {
            { TFAT.Column.Email, emailBox.Text },
            { TFAT.Column.Training, nameBox.Text },
            { TFAT.Column.Date, dateBox.Text },
            { TFAT.Column.First, fnBox.Text },
            { TFAT.Column.Last, lnBox.Text },
            { TFAT.Column.Rank, rankBox.Text },
        };
        public TFATSettings()
        {
            InitializeComponent();

            try
            {
                trainings = new BindingList<TFAT.TFATTraining>(TFAT.LoadTrainings("trainings.xml"));
            } catch
            {
                
            }

            if (trainings.Count == 0)
                trgAddBtn_Click(null, null);

            trgList.DataSource = trainings;
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
            Settings.Default.TFATEmailCol = emailBox.Text;
            Settings.Default.TFATTrgCol = nameBox.Text;
            Settings.Default.TFATDateCol =  dateBox.Text;
            Settings.Default.TFATFNameCol = fnBox.Text;
            Settings.Default.TFATLNameCol = lnBox.Text;
            Settings.Default.Save();

            TFAT.SaveTrainings("trainings.xml", Trainings());
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
    }
}