namespace Antal_PC
{
    partial class Form1
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
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btn_file = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.btn_PC_file = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(498, 222);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(364, 159);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(234, 20);
            this.textBox1.TabIndex = 1;
            // 
            // btn_file
            // 
            this.btn_file.Location = new System.Drawing.Point(605, 159);
            this.btn_file.Name = "btn_file";
            this.btn_file.Size = new System.Drawing.Size(26, 20);
            this.btn_file.TabIndex = 2;
            this.btn_file.Text = "...";
            this.btn_file.UseVisualStyleBackColor = true;
            this.btn_file.Click += new System.EventHandler(this.btn_file_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(364, 186);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(234, 20);
            this.textBox2.TabIndex = 3;
            // 
            // btn_PC_file
            // 
            this.btn_PC_file.Location = new System.Drawing.Point(605, 186);
            this.btn_PC_file.Name = "btn_PC_file";
            this.btn_PC_file.Size = new System.Drawing.Size(26, 20);
            this.btn_PC_file.TabIndex = 4;
            this.btn_PC_file.Text = "...";
            this.btn_PC_file.UseVisualStyleBackColor = true;
            this.btn_PC_file.Click += new System.EventHandler(this.btn_PC_file_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btn_PC_file);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.btn_file);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btn_file;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button btn_PC_file;
    }
}

