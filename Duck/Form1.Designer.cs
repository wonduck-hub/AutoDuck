namespace Duck
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            menuStrip1 = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            openFileToolStripMenuItem = new ToolStripMenuItem();
            showExcelWindowCheckBox = new CheckBox();
            worksheetsComboBox = new ComboBox();
            runButton = new Button();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(800, 24);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { openFileToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(37, 20);
            fileToolStripMenuItem.Text = "File";
            // 
            // openFileToolStripMenuItem
            // 
            openFileToolStripMenuItem.Name = "openFileToolStripMenuItem";
            openFileToolStripMenuItem.Size = new Size(180, 22);
            openFileToolStripMenuItem.Text = "Open";
            openFileToolStripMenuItem.Click += openFileToolStripMenuItem_Click;
            // 
            // showExcelWindowCheckBox
            // 
            showExcelWindowCheckBox.AutoSize = true;
            showExcelWindowCheckBox.Location = new Point(689, 47);
            showExcelWindowCheckBox.Name = "showExcelWindowCheckBox";
            showExcelWindowCheckBox.Size = new Size(99, 19);
            showExcelWindowCheckBox.TabIndex = 2;
            showExcelWindowCheckBox.Text = "Excel Window";
            showExcelWindowCheckBox.UseVisualStyleBackColor = true;
            showExcelWindowCheckBox.CheckedChanged += showExcelWindowCheckBox_CheckedChanged;
            // 
            // worksheetsComboBox
            // 
            worksheetsComboBox.FormattingEnabled = true;
            worksheetsComboBox.Location = new Point(12, 47);
            worksheetsComboBox.Name = "worksheetsComboBox";
            worksheetsComboBox.Size = new Size(121, 23);
            worksheetsComboBox.TabIndex = 3;
            // 
            // runButton
            // 
            runButton.Location = new Point(563, 129);
            runButton.Name = "runButton";
            runButton.Size = new Size(225, 127);
            runButton.TabIndex = 4;
            runButton.Text = "Run";
            runButton.UseVisualStyleBackColor = true;
            runButton.Click += runButton_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(runButton);
            Controls.Add(worksheetsComboBox);
            Controls.Add(showExcelWindowCheckBox);
            Controls.Add(menuStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            Name = "Form1";
            Text = "Duck";
            FormClosed += form1_Close;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem openFileToolStripMenuItem;
        private CheckBox showExcelWindowCheckBox;
        private ComboBox worksheetsComboBox;
        private Button runButton;
    }
}
