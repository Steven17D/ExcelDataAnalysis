using System.Windows.Forms;
namespace Excel
{
    partial class Menu
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Menu));
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.openFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.calculateSummeryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.pathLable = new System.Windows.Forms.Label();
            this.sheetSelect = new System.Windows.Forms.ComboBox();
            this.toolStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip
            // 
            this.toolStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripDropDownButton1});
            this.toolStrip.Location = new System.Drawing.Point(0, 0);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Size = new System.Drawing.Size(484, 25);
            this.toolStrip.TabIndex = 1;
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openFileToolStripMenuItem,
            this.calculateSummeryToolStripMenuItem});
            this.toolStripDropDownButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton1.Image")));
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(41, 22);
            this.toolStripDropDownButton1.Text = "FILE";
            // 
            // openFileToolStripMenuItem
            // 
            this.openFileToolStripMenuItem.Name = "openFileToolStripMenuItem";
            this.openFileToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.openFileToolStripMenuItem.Text = "Open File...";
            this.openFileToolStripMenuItem.Click += new System.EventHandler(this.openFileToolStripMenuItem_Click);
            // 
            // calculateSummeryToolStripMenuItem
            // 
            this.calculateSummeryToolStripMenuItem.Enabled = false;
            this.calculateSummeryToolStripMenuItem.Name = "calculateSummeryToolStripMenuItem";
            this.calculateSummeryToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.calculateSummeryToolStripMenuItem.Text = "Calculate Summery";
            this.calculateSummeryToolStripMenuItem.Click += new System.EventHandler(this.calculateSummeryToolStripMenuItem_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.AllowUserToResizeRows = false;
            this.dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(12, 45);
            this.dataGridView.MinimumSize = new System.Drawing.Size(460, 404);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(460, 404);
            this.dataGridView.TabIndex = 2;
            // 
            // pathLable
            // 
            this.pathLable.AutoSize = true;
            this.pathLable.Location = new System.Drawing.Point(12, 29);
            this.pathLable.Name = "pathLable";
            this.pathLable.Size = new System.Drawing.Size(0, 13);
            this.pathLable.TabIndex = 3;
            // 
            // sheetSelect
            // 
            this.sheetSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.sheetSelect.FormattingEnabled = true;
            this.sheetSelect.Location = new System.Drawing.Point(351, 21);
            this.sheetSelect.Name = "sheetSelect";
            this.sheetSelect.Size = new System.Drawing.Size(121, 21);
            this.sheetSelect.TabIndex = 4;
            this.sheetSelect.Visible = false;
            this.sheetSelect.SelectedValueChanged += new System.EventHandler(this.sheetSelect_SelectedValueChanged);
            this.sheetSelect.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.sheetSelect_MouseWheel);
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 461);
            this.Controls.Add(this.sheetSelect);
            this.Controls.Add(this.pathLable);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.toolStrip);
            this.MaximumSize = new System.Drawing.Size(500, 500);
            this.MinimumSize = new System.Drawing.Size(500, 500);
            this.Name = "Menu";
            this.Text = "Menu";
            this.Resize += new System.EventHandler(this.Menu_Resize);
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton1;
        private System.Windows.Forms.ToolStripMenuItem openFileToolStripMenuItem;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Label pathLable;
        private System.Windows.Forms.ComboBox sheetSelect;
        private ToolStripMenuItem calculateSummeryToolStripMenuItem;

    }
}