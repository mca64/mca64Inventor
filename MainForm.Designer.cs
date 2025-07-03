namespace mca64Inventor
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.DataGridView dataGridViewParts;
        private System.Windows.Forms.Button buttonGenerateThumbnails;
        private System.Windows.Forms.Button buttonGrawerowanie;
        private System.Windows.Forms.Label labelFontSize;
        private System.Windows.Forms.TextBox textBoxFontSize;
        private System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.CheckBox checkBoxCloseParts;
        private System.Windows.Forms.ComboBox comboBoxFonts;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.dataGridViewParts = new System.Windows.Forms.DataGridView();
            this.buttonGenerateThumbnails = new System.Windows.Forms.Button();
            this.buttonGrawerowanie = new System.Windows.Forms.Button();
            this.labelFontSize = new System.Windows.Forms.Label();
            this.textBoxFontSize = new System.Windows.Forms.TextBox();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.checkBoxCloseParts = new System.Windows.Forms.CheckBox();
            this.comboBoxFonts = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewParts)).BeginInit();
            this.SuspendLayout();
            // 
            // Controls layout in a single row, evenly distributed across the form width
            int margin = 10;
            int controlsCount = 6; // labelFontSize, textBoxFontSize, comboBoxFonts, checkBoxCloseParts, buttonGrawerowanie, buttonGenerateThumbnails
            int formWidth = this.ClientSize.Width;
            int controlWidth = (formWidth - (controlsCount + 1) * margin) / controlsCount;
            int controlHeight = 30;
            int y = 10;

            this.labelFontSize.Location = new System.Drawing.Point(margin, y);
            this.labelFontSize.Size = new System.Drawing.Size(controlWidth, controlHeight);
            this.labelFontSize.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.labelFontSize.Text = "Font size:";

            this.textBoxFontSize.Location = new System.Drawing.Point(margin + (controlWidth + margin) * 1, y);
            this.textBoxFontSize.Size = new System.Drawing.Size(controlWidth, controlHeight);
            this.textBoxFontSize.Text = "2.0";

            this.comboBoxFonts.Location = new System.Drawing.Point(margin + (controlWidth + margin) * 2, y);
            this.comboBoxFonts.Size = new System.Drawing.Size(controlWidth, controlHeight);

            this.checkBoxCloseParts.Location = new System.Drawing.Point(margin + (controlWidth + margin) * 3, y);
            this.checkBoxCloseParts.Size = new System.Drawing.Size(controlWidth, controlHeight);
            this.checkBoxCloseParts.Text = "Close parts after engraving";
            this.checkBoxCloseParts.Checked = true;
            this.checkBoxCloseParts.TabIndex = 10;

            this.buttonGrawerowanie.Location = new System.Drawing.Point(margin + (controlWidth + margin) * 4, y);
            this.buttonGrawerowanie.Size = new System.Drawing.Size(controlWidth, controlHeight);
            this.buttonGrawerowanie.TabIndex = 1;
            this.buttonGrawerowanie.Text = "Engrave";
            this.buttonGrawerowanie.UseVisualStyleBackColor = true;
            this.buttonGrawerowanie.Click += new System.EventHandler(this.buttonGrawerowanie_Click);

            this.buttonGenerateThumbnails.Location = new System.Drawing.Point(margin + (controlWidth + margin) * 5, y);
            this.buttonGenerateThumbnails.Size = new System.Drawing.Size(controlWidth, controlHeight);
            this.buttonGenerateThumbnails.TabIndex = 4;
            this.buttonGenerateThumbnails.Text = "Thumbnails";
            this.buttonGenerateThumbnails.UseVisualStyleBackColor = true;
            this.buttonGenerateThumbnails.Click += new System.EventHandler(this.buttonGenerateThumbnails_Click);

            // Add resize handler to keep controls evenly distributed
            this.Resize += (s, e) =>
            {
                int w = this.ClientSize.Width;
                int cw = (w - (controlsCount + 1) * margin) / controlsCount;
                this.labelFontSize.Location = new System.Drawing.Point(margin, y);
                this.labelFontSize.Size = new System.Drawing.Size(cw, controlHeight);
                this.textBoxFontSize.Location = new System.Drawing.Point(margin + (cw + margin) * 1, y);
                this.textBoxFontSize.Size = new System.Drawing.Size(cw, controlHeight);
                this.comboBoxFonts.Location = new System.Drawing.Point(margin + (cw + margin) * 2, y);
                this.comboBoxFonts.Size = new System.Drawing.Size(cw, controlHeight);
                this.checkBoxCloseParts.Location = new System.Drawing.Point(margin + (cw + margin) * 3, y);
                this.checkBoxCloseParts.Size = new System.Drawing.Size(cw, controlHeight);
                this.buttonGrawerowanie.Location = new System.Drawing.Point(margin + (cw + margin) * 4, y);
                this.buttonGrawerowanie.Size = new System.Drawing.Size(cw, controlHeight);
                this.buttonGenerateThumbnails.Location = new System.Drawing.Point(margin + (cw + margin) * 5, y);
                this.buttonGenerateThumbnails.Size = new System.Drawing.Size(cw, controlHeight);
            };
            // 
            // dataGridViewParts
            // 
            this.dataGridViewParts.AllowUserToAddRows = false;
            this.dataGridViewParts.AllowUserToDeleteRows = false;
            this.dataGridViewParts.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewParts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewParts.Location = new System.Drawing.Point(20, 50);
            this.dataGridViewParts.Name = "dataGridViewParts";
            this.dataGridViewParts.ReadOnly = false;
            this.dataGridViewParts.RowTemplate.Height = 64;
            this.dataGridViewParts.Size = new System.Drawing.Size(760, 280);
            this.dataGridViewParts.TabIndex = 2;
            // 
            // textBoxLog
            // 
            this.textBoxLog.Location = new System.Drawing.Point(20, 340);
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.Size = new System.Drawing.Size(760, 100);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxLog.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxLog.BringToFront();
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 460);
            this.Controls.Add(this.labelFontSize);
            this.Controls.Add(this.textBoxFontSize);
            this.Controls.Add(this.comboBoxFonts);
            this.Controls.Add(this.checkBoxCloseParts);
            this.Controls.Add(this.buttonGrawerowanie);
            this.Controls.Add(this.buttonGenerateThumbnails);
            this.Controls.Add(this.dataGridViewParts);
            this.Controls.Add(this.textBoxLog);
            this.Name = "MainForm";
            this.Text = "mca64launcher";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewParts)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
