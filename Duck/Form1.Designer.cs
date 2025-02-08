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
            proteinInfoButton = new Button();
            proteinInfoTextBox = new TextBox();
            menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)percentageNumericUpDown).BeginInit();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new Size(36, 36);
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Padding = new Padding(13, 5, 0, 5);
            menuStrip1.Size = new Size(1714, 51);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { openFileToolStripMenuItem, saveFileToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(80, 41);
            fileToolStripMenuItem.Text = "File";
            // 
            // openFileToolStripMenuItem
            // 
            openFileToolStripMenuItem.Name = "openFileToolStripMenuItem";
            openFileToolStripMenuItem.Size = new Size(222, 48);
            openFileToolStripMenuItem.Text = "File";
            openFileToolStripMenuItem.Click += openFileToolStripMenuItem_Click;
            // 
            // saveFileToolStripMenuItem
            // 
            saveFileToolStripMenuItem.Name = "saveFileToolStripMenuItem";
            saveFileToolStripMenuItem.Size = new Size(222, 48);
            saveFileToolStripMenuItem.Text = "Save";
            saveFileToolStripMenuItem.Click += saveToolStripMenuItem_Click;
            // 
            // showExcelWindowCheckBox
            // 
            showExcelWindowCheckBox.AutoSize = true;
            showExcelWindowCheckBox.Location = new Point(1476, 123);
            showExcelWindowCheckBox.Margin = new Padding(6, 7, 6, 7);
            showExcelWindowCheckBox.Name = "showExcelWindowCheckBox";
            showExcelWindowCheckBox.Size = new Size(214, 41);
            showExcelWindowCheckBox.TabIndex = 2;
            showExcelWindowCheckBox.Text = "Excel Window";
            showExcelWindowCheckBox.UseVisualStyleBackColor = true;
            showExcelWindowCheckBox.CheckedChanged += showExcelWindowCheckBox_CheckedChanged;
            // 
            // worksheetsComboBox
            // 
            worksheetsComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            worksheetsComboBox.FormattingEnabled = true;
            worksheetsComboBox.Location = new Point(315, 118);
            worksheetsComboBox.Margin = new Padding(6, 7, 6, 7);
            worksheetsComboBox.Name = "worksheetsComboBox";
            worksheetsComboBox.Size = new Size(255, 45);
            worksheetsComboBox.TabIndex = 3;
            // 
            // runButton
            // 
            runButton.Location = new Point(26, 272);
            runButton.Margin = new Padding(6, 7, 6, 7);
            runButton.Name = "runButton";
            runButton.Size = new Size(266, 113);
            runButton.TabIndex = 4;
            runButton.Text = "Run";
            runButton.UseVisualStyleBackColor = true;
            runButton.Click += runButton_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(26, 126);
            label1.Margin = new Padding(6, 0, 6, 0);
            label1.Name = "label1";
            label1.Size = new Size(190, 37);
            label1.TabIndex = 5;
            label1.Text = "selected sheet:";
            // 
            // percentageNumericUpDown
            // 
            percentageNumericUpDown.DecimalPlaces = 2;
            percentageNumericUpDown.Increment = new decimal(new int[] { 1, 0, 0, 65536 });
            percentageNumericUpDown.Location = new Point(315, 190);
            percentageNumericUpDown.Margin = new Padding(6, 7, 6, 7);
            percentageNumericUpDown.Maximum = new decimal(new int[] { 1, 0, 0, 0 });
            percentageNumericUpDown.Name = "percentageNumericUpDown";
            percentageNumericUpDown.Size = new Size(257, 43);
            percentageNumericUpDown.TabIndex = 6;
            percentageNumericUpDown.Value = new decimal(new int[] { 1, 0, 0, 65536 });
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(26, 195);
            label2.Margin = new Padding(6, 0, 6, 0);
            label2.Name = "label2";
            label2.Size = new Size(280, 37);
            label2.TabIndex = 7;
            label2.Text = "extraction percentage:";
            // 
            // proteinInfoButton
            // 
            proteinInfoButton.Location = new Point(26, 395);
            proteinInfoButton.Margin = new Padding(6, 7, 6, 7);
            proteinInfoButton.Name = "proteinInfoButton";
            proteinInfoButton.Size = new Size(266, 113);
            proteinInfoButton.TabIndex = 8;
            proteinInfoButton.Text = "Protein info";
            proteinInfoButton.UseVisualStyleBackColor = true;
            proteinInfoButton.Click += proteinInfoButton_Click;
            // 
            // proteinInfoTextBox
            // 
            proteinInfoTextBox.Location = new Point(315, 395);
            proteinInfoTextBox.Margin = new Padding(6, 7, 6, 7);
            proteinInfoTextBox.Multiline = true;
            proteinInfoTextBox.Name = "proteinInfoTextBox";
            proteinInfoTextBox.ReadOnly = true;
            proteinInfoTextBox.ScrollBars = ScrollBars.Vertical;
            proteinInfoTextBox.Size = new Size(1369, 699);
            proteinInfoTextBox.TabIndex = 10;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(15F, 37F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1714, 1110);
            Controls.Add(proteinInfoTextBox);
            Controls.Add(proteinInfoButton);
            Controls.Add(label2);
            Controls.Add(percentageNumericUpDown);
            Controls.Add(label1);
            Controls.Add(runButton);
            Controls.Add(worksheetsComboBox);
            Controls.Add(showExcelWindowCheckBox);
            Controls.Add(menuStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            Margin = new Padding(6, 7, 6, 7);
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
        private Button proteinInfoButton;
        private TextBox proteinInfoTextBox;
    }
}
