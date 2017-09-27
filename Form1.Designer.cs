namespace PARSEEXCEL
{
    partial class mainForm
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
            this.mainComboBox = new System.Windows.Forms.ComboBox();
            this.mainTextBox = new System.Windows.Forms.RichTextBox();
            this.startButton = new System.Windows.Forms.Button();
            this.mainProgressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // mainComboBox
            // 
            this.mainComboBox.FormattingEnabled = true;
            this.mainComboBox.Location = new System.Drawing.Point(12, 12);
            this.mainComboBox.Name = "mainComboBox";
            this.mainComboBox.Size = new System.Drawing.Size(131, 21);
            this.mainComboBox.TabIndex = 0;
            this.mainComboBox.SelectedIndexChanged += new System.EventHandler(this.mainComboBox_SelectedIndexChanged);
            // 
            // mainTextBox
            // 
            this.mainTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.mainTextBox.Location = new System.Drawing.Point(149, 12);
            this.mainTextBox.Name = "mainTextBox";
            this.mainTextBox.Size = new System.Drawing.Size(631, 317);
            this.mainTextBox.TabIndex = 1;
            this.mainTextBox.Text = "";
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(13, 40);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(130, 23);
            this.startButton.TabIndex = 2;
            this.startButton.Text = "Получить список групп";
            this.startButton.UseVisualStyleBackColor = true;
            this.startButton.Click += new System.EventHandler(this.startButton_Click);
            // 
            // mainProgressBar
            // 
            this.mainProgressBar.Location = new System.Drawing.Point(149, 335);
            this.mainProgressBar.Name = "mainProgressBar";
            this.mainProgressBar.Size = new System.Drawing.Size(631, 23);
            this.mainProgressBar.TabIndex = 3;
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(792, 370);
            this.Controls.Add(this.mainProgressBar);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.mainTextBox);
            this.Controls.Add(this.mainComboBox);
            this.MinimumSize = new System.Drawing.Size(400, 200);
            this.Name = "mainForm";
            this.Text = "Schedule";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox mainComboBox;
        private System.Windows.Forms.RichTextBox mainTextBox;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.ProgressBar mainProgressBar;
    }
}

