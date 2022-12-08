namespace MyDepression
{
    partial class TFATAssociation
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lnBox = new System.Windows.Forms.TextBox();
            this.fnBox = new System.Windows.Forms.TextBox();
            this.dateBox = new System.Windows.Forms.TextBox();
            this.nameBox = new System.Windows.Forms.TextBox();
            this.emailBox = new System.Windows.Forms.TextBox();
            this.rankBox = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(227, 198);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(115, 37);
            this.button1.TabIndex = 0;
            this.button1.Text = "Apply";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(32, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "Email:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 20);
            this.label2.TabIndex = 5;
            this.label2.Text = "Training:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(37, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 20);
            this.label3.TabIndex = 6;
            this.label3.Text = "Date:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(39, 105);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(50, 20);
            this.label4.TabIndex = 9;
            this.label4.Text = "First:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(40, 137);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 20);
            this.label5.TabIndex = 10;
            this.label5.Text = "Last:";
            // 
            // lnBox
            // 
            this.lnBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::MyDepression.Properties.Settings.Default, "TFATLNameCol", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.lnBox.Location = new System.Drawing.Point(95, 134);
            this.lnBox.Name = "lnBox";
            this.lnBox.Size = new System.Drawing.Size(247, 26);
            this.lnBox.TabIndex = 8;
            this.lnBox.Text = global::MyDepression.Properties.Settings.Default.TFATLNameCol;
            // 
            // fnBox
            // 
            this.fnBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::MyDepression.Properties.Settings.Default, "TFATFNameCol", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.fnBox.Location = new System.Drawing.Point(95, 102);
            this.fnBox.Name = "fnBox";
            this.fnBox.Size = new System.Drawing.Size(247, 26);
            this.fnBox.TabIndex = 7;
            this.fnBox.Text = global::MyDepression.Properties.Settings.Default.TFATFNameCol;
            // 
            // dateBox
            // 
            this.dateBox.Location = new System.Drawing.Point(96, 70);
            this.dateBox.Name = "dateBox";
            this.dateBox.Size = new System.Drawing.Size(247, 26);
            this.dateBox.TabIndex = 4;
            this.dateBox.Text = global::MyDepression.Properties.Settings.Default.TFATDateCol;
            // 
            // nameBox
            // 
            this.nameBox.Location = new System.Drawing.Point(96, 38);
            this.nameBox.Name = "nameBox";
            this.nameBox.Size = new System.Drawing.Size(247, 26);
            this.nameBox.TabIndex = 3;
            this.nameBox.Text = global::MyDepression.Properties.Settings.Default.TFATTrgCol;
            // 
            // emailBox
            // 
            this.emailBox.Location = new System.Drawing.Point(96, 6);
            this.emailBox.Name = "emailBox";
            this.emailBox.Size = new System.Drawing.Size(247, 26);
            this.emailBox.TabIndex = 2;
            this.emailBox.Text = global::MyDepression.Properties.Settings.Default.TFATEmailCol;
            // 
            // rankBox
            // 
            this.rankBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::MyDepression.Properties.Settings.Default, "TFATRankCol", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.rankBox.Location = new System.Drawing.Point(95, 166);
            this.rankBox.Name = "rankBox";
            this.rankBox.Size = new System.Drawing.Size(247, 26);
            this.rankBox.TabIndex = 11;
            this.rankBox.Text = global::MyDepression.Properties.Settings.Default.TFATRankCol;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(32, 169);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(56, 20);
            this.label6.TabIndex = 12;
            this.label6.Text = "Rank:";
            // 
            // TFATAssociation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(354, 252);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.rankBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lnBox);
            this.Controls.Add(this.fnBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dateBox);
            this.Controls.Add(this.nameBox);
            this.Controls.Add(this.emailBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "TFATAssociation";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "TFAT Associations";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.TFATAssociation_FormClosing);
            this.Load += new System.EventHandler(this.TFATAssociation_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox emailBox;
        private System.Windows.Forms.TextBox nameBox;
        private System.Windows.Forms.TextBox dateBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox fnBox;
        private System.Windows.Forms.TextBox lnBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox rankBox;
        private System.Windows.Forms.Label label6;
    }
}