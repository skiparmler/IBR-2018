namespace SKIIBR
{
    partial class Form1
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
            this.btnExecute = new System.Windows.Forms.Button();
            this.numYear = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.cbBranscher = new System.Windows.Forms.ComboBox();
            this.cbSegment = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.lblPath = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.tbOrt = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.dgvResultatVariabler = new System.Windows.Forms.DataGridView();
            this.dgvResultatVariablerNamn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvResultatVariablerText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvKundDimensioner = new System.Windows.Forms.DataGridView();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.dataGridViewKundDimensionerNamn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewKundDimensionerText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.numYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResultatVariabler)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvKundDimensioner)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(138, 482);
            this.btnExecute.Margin = new System.Windows.Forms.Padding(2);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(56, 19);
            this.btnExecute.TabIndex = 0;
            this.btnExecute.Text = "Kör";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // numYear
            // 
            this.numYear.Location = new System.Drawing.Point(83, 369);
            this.numYear.Margin = new System.Windows.Forms.Padding(2);
            this.numYear.Maximum = new decimal(new int[] {
            2100,
            0,
            0,
            0});
            this.numYear.Name = "numYear";
            this.numYear.Size = new System.Drawing.Size(90, 20);
            this.numYear.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 328);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Bransch:";
            // 
            // cbBranscher
            // 
            this.cbBranscher.FormattingEnabled = true;
            this.cbBranscher.Location = new System.Drawing.Point(83, 322);
            this.cbBranscher.Margin = new System.Windows.Forms.Padding(2);
            this.cbBranscher.Name = "cbBranscher";
            this.cbBranscher.Size = new System.Drawing.Size(92, 21);
            this.cbBranscher.TabIndex = 5;
            // 
            // cbSegment
            // 
            this.cbSegment.FormattingEnabled = true;
            this.cbSegment.Location = new System.Drawing.Point(83, 344);
            this.cbSegment.Margin = new System.Windows.Forms.Padding(2);
            this.cbSegment.Name = "cbSegment";
            this.cbSegment.Size = new System.Drawing.Size(92, 21);
            this.cbSegment.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 349);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Segment:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 372);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(20, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "År:";
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(27, 11);
            this.btnSelectFolder.Margin = new System.Windows.Forms.Padding(2);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(90, 22);
            this.btnSelectFolder.TabIndex = 9;
            this.btnSelectFolder.Text = "Välj mapp";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(157, 16);
            this.lblPath.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(0, 13);
            this.lblPath.TabIndex = 10;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(24, 395);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Datum:";
            // 
            // dateTimePicker
            // 
            this.dateTimePicker.Location = new System.Drawing.Point(83, 395);
            this.dateTimePicker.Name = "dateTimePicker";
            this.dateTimePicker.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker.TabIndex = 12;
            // 
            // tbOrt
            // 
            this.tbOrt.Location = new System.Drawing.Point(83, 422);
            this.tbOrt.Name = "tbOrt";
            this.tbOrt.Size = new System.Drawing.Size(100, 20);
            this.tbOrt.TabIndex = 13;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(24, 425);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(24, 13);
            this.label5.TabIndex = 14;
            this.label5.Text = "Ort:";
            // 
            // dgvResultatVariabler
            // 
            this.dgvResultatVariabler.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvResultatVariabler.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dgvResultatVariablerNamn,
            this.dgvResultatVariablerText});
            this.dgvResultatVariabler.Location = new System.Drawing.Point(27, 62);
            this.dgvResultatVariabler.Name = "dgvResultatVariabler";
            this.dgvResultatVariabler.Size = new System.Drawing.Size(256, 109);
            this.dgvResultatVariabler.TabIndex = 16;
            // 
            // dgvResultatVariablerNamn
            // 
            this.dgvResultatVariablerNamn.HeaderText = "Namn";
            this.dgvResultatVariablerNamn.Name = "dgvResultatVariablerNamn";
            // 
            // dgvResultatVariablerText
            // 
            this.dgvResultatVariablerText.HeaderText = "Text";
            this.dgvResultatVariablerText.Name = "dgvResultatVariablerText";
            // 
            // dgvKundDimensioner
            // 
            this.dgvKundDimensioner.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvKundDimensioner.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewKundDimensionerNamn,
            this.dataGridViewKundDimensionerText});
            this.dgvKundDimensioner.Location = new System.Drawing.Point(27, 198);
            this.dgvKundDimensioner.Name = "dgvKundDimensioner";
            this.dgvKundDimensioner.Size = new System.Drawing.Size(256, 109);
            this.dgvKundDimensioner.TabIndex = 16;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(28, 46);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "Resultatvariabler:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(28, 182);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(91, 13);
            this.label7.TabIndex = 17;
            this.label7.Text = "Kunddimensioner:";
            // 
            // dataGridViewKundDimensionerNamn
            // 
            this.dataGridViewKundDimensionerNamn.HeaderText = "Namn";
            this.dataGridViewKundDimensionerNamn.Name = "dataGridViewKundDimensionerNamn";
            // 
            // dataGridViewKundDimensionerText
            // 
            this.dataGridViewKundDimensionerText.HeaderText = "Text";
            this.dataGridViewKundDimensionerText.Name = "dataGridViewKundDimensionerText";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(334, 514);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.dgvKundDimensioner);
            this.Controls.Add(this.dgvResultatVariabler);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.tbOrt);
            this.Controls.Add(this.dateTimePicker);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cbSegment);
            this.Controls.Add(this.cbBranscher);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.numYear);
            this.Controls.Add(this.btnExecute);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "IBR";
            ((System.ComponentModel.ISupportInitialize)(this.numYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResultatVariabler)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvKundDimensioner)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExecute;
        private System.Windows.Forms.NumericUpDown numYear;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbBranscher;
        private System.Windows.Forms.ComboBox cbSegment;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.FolderBrowserDialog folderBrowser;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker dateTimePicker;
        private System.Windows.Forms.TextBox tbOrt;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DataGridView dgvResultatVariabler;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvResultatVariablerNamn;
        private System.Windows.Forms.DataGridViewTextBoxColumn dgvResultatVariablerText;
        private System.Windows.Forms.DataGridView dgvKundDimensioner;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewKundDimensionerNamn;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewKundDimensionerText;
    }
}

