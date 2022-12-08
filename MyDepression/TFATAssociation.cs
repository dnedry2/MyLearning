using MyDepression.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyDepression
{
    public partial class TFATAssociation : Form
    {
        public Dictionary<TFAT.Column, string> Associations() => new Dictionary<TFAT.Column, string>()
        {
            { TFAT.Column.Email, emailBox.Text },
            { TFAT.Column.Training, nameBox.Text },
            { TFAT.Column.Date, dateBox.Text },
            { TFAT.Column.First, fnBox.Text },
            { TFAT.Column.Last, lnBox.Text },
            { TFAT.Column.Rank, rankBox.Text },
        };
        public TFATAssociation()
        {
            InitializeComponent();
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
        }
    }
}