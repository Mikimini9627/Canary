namespace Canary
{
    partial class ExcelViewer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelViewer));
            reoGridControl = new unvell.ReoGrid.ReoGridControl();
            SuspendLayout();
            // 
            // reoGridControl
            // 
            reoGridControl.BackColor = Color.FromArgb(255, 255, 255);
            reoGridControl.ColumnHeaderContextMenuStrip = null;
            reoGridControl.Dock = DockStyle.Fill;
            reoGridControl.LeadHeaderContextMenuStrip = null;
            reoGridControl.Location = new Point(0, 0);
            reoGridControl.Name = "reoGridControl";
            reoGridControl.RowHeaderContextMenuStrip = null;
            reoGridControl.Script = null;
            reoGridControl.SheetTabContextMenuStrip = null;
            reoGridControl.SheetTabNewButtonVisible = true;
            reoGridControl.SheetTabVisible = true;
            reoGridControl.SheetTabWidth = 60;
            reoGridControl.ShowScrollEndSpacing = true;
            reoGridControl.Size = new Size(1044, 562);
            reoGridControl.TabIndex = 0;
            reoGridControl.Text = "reoGridControl1";
            // 
            // ExcelViewer
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1044, 562);
            Controls.Add(reoGridControl);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "ExcelViewer";
            StartPosition = FormStartPosition.CenterParent;
            Text = "検索結果一覧";
            ResumeLayout(false);
        }

        #endregion

        private unvell.ReoGrid.ReoGridControl reoGridControl;
    }
}