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
    public partial class ExceptionForm : Form
    {
        public ExceptionForm(Exception ex, string msg = "An error occured:")
        {
            InitializeComponent();

            textBox1.Text = ex.Message;
            textBox2.Text = ex.StackTrace;

            label1.Text = msg;
        }
    }
}