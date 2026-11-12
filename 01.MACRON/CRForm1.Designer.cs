namespace CRPRT
{
    partial class CRForm1
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
            this.CrView1 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.SuspendLayout();
            // 
            // CrView1
            // 
            this.CrView1.AccessibleRole = System.Windows.Forms.AccessibleRole.Window;
            this.CrView1.ActiveViewIndex = -1;
            this.CrView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.CrView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CrView1.Location = new System.Drawing.Point(0, 0);
            this.CrView1.Name = "CrView1";
            this.CrView1.SelectionFormula = "";
            this.CrView1.Size = new System.Drawing.Size(528, 398);
            this.CrView1.TabIndex = 0;
            this.CrView1.ViewTimeSelectionFormula = "";
            // 
            // CRForm1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 398);
            this.Controls.Add(this.CrView1);
            this.Name = "CRForm1";
            this.Text = "CRForm1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.ResumeLayout(false);

        }

        #endregion

        public CrystalDecisions.Windows.Forms.CrystalReportViewer CrView1;

    }
}