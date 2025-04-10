namespace ISTools
{
    partial class ParamCombineForm
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
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.discriptionTextBox1 = new System.Windows.Forms.TextBox();
            this.inputTextBox1 = new System.Windows.Forms.TextBox();
            this.discriptionTextBox2 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // comboBox1
            // 
            this.comboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(275, 23);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(493, 24);
            this.comboBox1.TabIndex = 0;
            // 
            // discriptionTextBox1
            // 
            this.discriptionTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.discriptionTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.discriptionTextBox1.Location = new System.Drawing.Point(12, 26);
            this.discriptionTextBox1.Name = "discriptionTextBox1";
            this.discriptionTextBox1.ReadOnly = true;
            this.discriptionTextBox1.Size = new System.Drawing.Size(250, 15);
            this.discriptionTextBox1.TabIndex = 1;
            // 
            // inputTextBox1
            // 
            this.inputTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.inputTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.inputTextBox1.Location = new System.Drawing.Point(24, 121);
            this.inputTextBox1.Name = "inputTextBox1";
            this.inputTextBox1.Size = new System.Drawing.Size(744, 22);
            this.inputTextBox1.TabIndex = 1;
            // 
            // discriptionTextBox2
            // 
            this.discriptionTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.discriptionTextBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.discriptionTextBox2.Location = new System.Drawing.Point(24, 57);
            this.discriptionTextBox2.Multiline = true;
            this.discriptionTextBox2.Name = "discriptionTextBox2";
            this.discriptionTextBox2.ReadOnly = true;
            this.discriptionTextBox2.Size = new System.Drawing.Size(744, 58);
            this.discriptionTextBox2.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(256, 156);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(254, 32);
            this.button1.TabIndex = 2;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 204);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(790, 14);
            this.progressBar1.TabIndex = 3;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(790, 218);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.inputTextBox1);
            this.Controls.Add(this.discriptionTextBox2);
            this.Controls.Add(this.discriptionTextBox1);
            this.Controls.Add(this.comboBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ComboBox comboBox1;
        public System.Windows.Forms.TextBox discriptionTextBox1;
        public System.Windows.Forms.TextBox inputTextBox1;
        public System.Windows.Forms.TextBox discriptionTextBox2;
        public System.Windows.Forms.Button button1;
        public System.Windows.Forms.ProgressBar progressBar1;
    }
}