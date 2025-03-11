namespace Qr_Generator
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtBoxXLocation = new System.Windows.Forms.TextBox();
            this.BtnBrowse = new System.Windows.Forms.Button();
            this.BtnDownloadTemplate = new System.Windows.Forms.Button();
            this.BtnConvert = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(342, 28);
            this.label1.TabIndex = 3;
            this.label1.Text = "Load the Excel Template with Values";
            // 
            // txtBoxXLocation
            // 
            this.txtBoxXLocation.Location = new System.Drawing.Point(18, 71);
            this.txtBoxXLocation.Name = "txtBoxXLocation";
            this.txtBoxXLocation.Size = new System.Drawing.Size(461, 22);
            this.txtBoxXLocation.TabIndex = 4;
            // 
            // BtnBrowse
            // 
            this.BtnBrowse.Location = new System.Drawing.Point(504, 67);
            this.BtnBrowse.Name = "BtnBrowse";
            this.BtnBrowse.Size = new System.Drawing.Size(96, 28);
            this.BtnBrowse.TabIndex = 5;
            this.BtnBrowse.Text = "Browse";
            this.BtnBrowse.UseVisualStyleBackColor = true;
            this.BtnBrowse.Click += new System.EventHandler(this.BtnBrowse_Click);
            // 
            // BtnDownloadTemplate
            // 
            this.BtnDownloadTemplate.Location = new System.Drawing.Point(306, 200);
            this.BtnDownloadTemplate.Name = "BtnDownloadTemplate";
            this.BtnDownloadTemplate.Size = new System.Drawing.Size(160, 37);
            this.BtnDownloadTemplate.TabIndex = 6;
            this.BtnDownloadTemplate.Text = "Download Template";
            this.BtnDownloadTemplate.UseVisualStyleBackColor = true;
            this.BtnDownloadTemplate.Click += new System.EventHandler(this.BtnDownloadTemplate_Click);
            // 
            // BtnConvert
            // 
            this.BtnConvert.Location = new System.Drawing.Point(472, 200);
            this.BtnConvert.Name = "BtnConvert";
            this.BtnConvert.Size = new System.Drawing.Size(132, 37);
            this.BtnConvert.TabIndex = 7;
            this.BtnConvert.Text = "Generate QR";
            this.BtnConvert.UseVisualStyleBackColor = true;
            this.BtnConvert.Click += new System.EventHandler(this.BtnConvert_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(4)))), ((int)(((byte)(59)))), ((int)(((byte)(91)))));
            this.ClientSize = new System.Drawing.Size(616, 253);
            this.Controls.Add(this.BtnConvert);
            this.Controls.Add(this.BtnDownloadTemplate);
            this.Controls.Add(this.BtnBrowse);
            this.Controls.Add(this.txtBoxXLocation);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "QR Generator V 1.3";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtBoxXLocation;
        private System.Windows.Forms.Button BtnBrowse;
        private System.Windows.Forms.Button BtnDownloadTemplate;
        private System.Windows.Forms.Button BtnConvert;
    }
}

