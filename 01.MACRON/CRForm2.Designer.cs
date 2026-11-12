namespace CRPRT
{
    partial class CRForm2
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
            this.CrView2 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.SuspendLayout();
            // 
            // CrView2
            // 
            this.CrView2.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.CrView2.ActiveViewIndex = -1;
            this.CrView2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.CrView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CrView2.Location = new System.Drawing.Point(0, 0);
            this.CrView2.Name = "CrView2";
            this.CrView2.SelectionFormula = "";
            this.CrView2.Size = new System.Drawing.Size(528, 398);
            this.CrView2.TabIndex = 0;
            this.CrView2.ViewTimeSelectionFormula = "";
            this.CrView2.Load += new System.EventHandler(this.CrView2_Load);
            // 
            // CRForm2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 398);
            this.Controls.Add(this.CrView2);
            this.Name = "CRForm2";
            this.Text = "CRForm2";
            this.TopMost = true;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.ResumeLayout(false);

        }

        #endregion

        public CrystalDecisions.Windows.Forms.CrystalReportViewer CrView2;

    }
}