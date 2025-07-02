namespace mca64Inventor
{
    partial class MainForm
    {
        /// <summary>
        /// Wymagana zmienna projektanta.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button buttonGrawerowanie;

        /// <summary>
        /// Wyczyœæ wszystkie u¿ywane zasoby.
        /// </summary>
        /// <param name="disposing">prawda, je¿eli zarz¹dzane zasoby powinny byæ zlikwidowane; w przeciwnym razie fa³sz.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Kod generowany przez Projektanta formularzy systemu Windows

        /// <summary>
        /// Metoda wymagana do obs³ugi projektanta — nie modyfikuj
        /// jej zawartoœci w edytorze kodu.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonGrawerowanie = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonGrawerowanie
            // 
            this.buttonGrawerowanie.Location = new System.Drawing.Point(30, 80);
            this.buttonGrawerowanie.Name = "buttonGrawerowanie";
            this.buttonGrawerowanie.Size = new System.Drawing.Size(200, 40);
            this.buttonGrawerowanie.TabIndex = 1;
            this.buttonGrawerowanie.Text = "Wykonaj grawerowanie";
            this.buttonGrawerowanie.UseVisualStyleBackColor = true;
            this.buttonGrawerowanie.Click += new System.EventHandler(this.buttonGrawerowanie_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 200);
            this.Controls.Add(this.buttonGrawerowanie);
            this.Name = "MainForm";
            this.Text = "mca64launcher";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
