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
            saveFileToolStripMenuItem = new ToolStripMenuItem();
            showExcelWindowCheckBox = new CheckBox();
            worksheetsComboBox = new ComboBox();
            runButton = new Button();
            label1 = new Label();
            percentageNumericUpDown = new NumericUpDown();
            label2 = new Label();
            menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)percentageNumericUpDown).BeginInit();
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
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { openFileToolStripMenuItem, saveFileToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(37, 20);
            fileToolStripMenuItem.Text = "File";
            // 
            // openFileToolStripMenuItem
            // 
            openFileToolStripMenuItem.Name = "openFileToolStripMenuItem";
            openFileToolStripMenuItem.Size = new Size(98, 22);
            openFileToolStripMenuItem.Text = "File";
            openFileToolStripMenuItem.Click += openFileToolStripMenuItem_Click;
            // 
            // saveFileToolStripMenuItem
            // 
            saveFileToolStripMenuItem.Name = "saveFileToolStripMenuItem";
            saveFileToolStripMenuItem.Size = new Size(98, 22);
            saveFileToolStripMenuItem.Text = "Save";
            saveFileToolStripMenuItem.Click += saveToolStripMenuItem_Click;
            // 
            // showExcelWindowCheckBox
            // 
            showExcelWindowCheckBox.AutoSize = true;
            showExcelWindowCheckBox.Location = new Point(689, 50);
            showExcelWindowCheckBox.Name = "showExcelWindowCheckBox";
            showExcelWindowCheckBox.Size = new Size(99, 19);
            showExcelWindowCheckBox.TabIndex = 2;
            showExcelWindowCheckBox.Text = "Excel Window";
            showExcelWindowCheckBox.UseVisualStyleBackColor = true;
            showExcelWindowCheckBox.CheckedChanged += showExcelWindowCheckBox_CheckedChanged;
            // 
            // worksheetsComboBox
            // 
            worksheetsComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            worksheetsComboBox.FormattingEnabled = true;
            worksheetsComboBox.Location = new Point(119, 48);
            worksheetsComboBox.Name = "worksheetsComboBox";
            worksheetsComboBox.Size = new Size(121, 23);
            worksheetsComboBox.TabIndex = 3;
            // 
            // runButton
            // 
            runButton.Location = new Point(12, 391);
            runButton.Name = "runButton";
            runButton.Size = new Size(776, 47);
            runButton.TabIndex = 4;
            runButton.Text = "Run";
            runButton.UseVisualStyleBackColor = true;
            runButton.Click += runButton_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 51);
            label1.Name = "label1";
            label1.Size = new Size(84, 15);
            label1.TabIndex = 5;
            label1.Text = "selected sheet:";
            // 
            // percentageNumericUpDown
            // 
            percentageNumericUpDown.DecimalPlaces = 2;
            percentageNumericUpDown.Increment = new decimal(new int[] { 1, 0, 0, 65536 });
            percentageNumericUpDown.Location = new Point(119, 77);
            percentageNumericUpDown.Maximum = new decimal(new int[] { 1, 0, 0, 0 });
            percentageNumericUpDown.Name = "percentageNumericUpDown";
            percentageNumericUpDown.Size = new Size(120, 23);
            percentageNumericUpDown.TabIndex = 6;
            percentageNumericUpDown.Value = new decimal(new int[] { 1, 0, 0, 65536 });
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 79);
            label2.Name = "label2";
            label2.Size = new Size(69, 15);
            label2.TabIndex = 7;
            label2.Text = "percentage:";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(label2);
            Controls.Add(percentageNumericUpDown);
            Controls.Add(label1);
            Controls.Add(runButton);
            Controls.Add(worksheetsComboBox);
            Controls.Add(showExcelWindowCheckBox);
            Controls.Add(menuStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            MaximumSize = new Size(816, 489);
            MinimumSize = new Size(816, 489);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "AutoDuck";
            FormClosed += form1_Close;
            Shown += form1_Shown;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)percentageNumericUpDown).EndInit();
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
        private ToolStripMenuItem saveFileToolStripMenuItem;
        private Label label1;
        private NumericUpDown percentageNumericUpDown;
        private Label label2;
    }
}
