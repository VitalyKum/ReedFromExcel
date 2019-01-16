namespace Postgen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.fdoExcel = new System.Windows.Forms.OpenFileDialog();
            this.txtChoose = new System.Windows.Forms.TextBox();
            this.cmdChoose = new System.Windows.Forms.Button();
            this.cmdOk = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.chkDoCorrect = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtReport = new System.Windows.Forms.TextBox();
            this.cmdSaveReport = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.panel1 = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Исходник";
            // 
            // fdoExcel
            // 
            this.fdoExcel.FileName = "openFileDialog1";
            this.fdoExcel.Filter = "\"Excel Files|*.xls;*.xlsx;*.xlsm\"";
            // 
            // txtChoose
            // 
            this.txtChoose.Location = new System.Drawing.Point(104, 42);
            this.txtChoose.Name = "txtChoose";
            this.txtChoose.Size = new System.Drawing.Size(335, 20);
            this.txtChoose.TabIndex = 1;
            // 
            // cmdChoose
            // 
            this.cmdChoose.Location = new System.Drawing.Point(445, 40);
            this.cmdChoose.Name = "cmdChoose";
            this.cmdChoose.Size = new System.Drawing.Size(24, 23);
            this.cmdChoose.TabIndex = 2;
            this.cmdChoose.Text = "...";
            this.cmdChoose.UseVisualStyleBackColor = true;
            this.cmdChoose.Click += new System.EventHandler(this.cmdChoose_Click);
            // 
            // cmdOk
            // 
            this.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.cmdOk.Location = new System.Drawing.Point(394, 73);
            this.cmdOk.Name = "cmdOk";
            this.cmdOk.Size = new System.Drawing.Size(74, 23);
            this.cmdOk.TabIndex = 3;
            this.cmdOk.Text = "Выполнить ";
            this.cmdOk.UseVisualStyleBackColor = true;
            this.cmdOk.Click += new System.EventHandler(this.cmdOk_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cmdCancel.Location = new System.Drawing.Point(394, 318);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 4;
            this.cmdCancel.Text = "Закрыть";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // chkDoCorrect
            // 
            this.chkDoCorrect.AutoSize = true;
            this.chkDoCorrect.Checked = true;
            this.chkDoCorrect.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDoCorrect.Location = new System.Drawing.Point(15, 79);
            this.chkDoCorrect.Name = "chkDoCorrect";
            this.chkDoCorrect.Size = new System.Drawing.Size(183, 17);
            this.chkDoCorrect.TabIndex = 5;
            this.chkDoCorrect.Text = "Корректировать по окончании ";
            this.chkDoCorrect.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 142);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Отчет об ошибках";
            // 
            // txtReport
            // 
            this.txtReport.ForeColor = System.Drawing.Color.RoyalBlue;
            this.txtReport.Location = new System.Drawing.Point(16, 158);
            this.txtReport.Multiline = true;
            this.txtReport.Name = "txtReport";
            this.txtReport.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtReport.Size = new System.Drawing.Size(453, 154);
            this.txtReport.TabIndex = 7;
            // 
            // cmdSaveReport
            // 
            this.cmdSaveReport.Location = new System.Drawing.Point(19, 319);
            this.cmdSaveReport.Name = "cmdSaveReport";
            this.cmdSaveReport.Size = new System.Drawing.Size(180, 23);
            this.cmdSaveReport.TabIndex = 8;
            this.cmdSaveReport.Text = "Сохранить отчет для отправки";
            this.cmdSaveReport.UseVisualStyleBackColor = true;
            this.cmdSaveReport.Click += new System.EventHandler(this.cmdSaveReport_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 105);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(455, 23);
            this.progressBar1.TabIndex = 9;
            // 
            // panel1
            // 
            this.panel1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("panel1.BackgroundImage")));
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(480, 37);
            this.panel1.TabIndex = 10;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(480, 350);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.cmdSaveReport);
            this.Controls.Add(this.txtReport);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chkDoCorrect);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdOk);
            this.Controls.Add(this.cmdChoose);
            this.Controls.Add(this.txtChoose);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "Подготовка файла почты v2.0 (Специально для Корона Мех)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog fdoExcel;
        private System.Windows.Forms.TextBox txtChoose;
        private System.Windows.Forms.Button cmdChoose;
        private System.Windows.Forms.Button cmdOk;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.CheckBox chkDoCorrect;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtReport;
        private System.Windows.Forms.Button cmdSaveReport;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Panel panel1;
    }
}

