namespace DotNetRefEdit
{
    partial class RefEditForm
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
            this.InputBox1 = new System.Windows.Forms.RichTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.InputBox2 = new System.Windows.Forms.RichTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.DestinationBox = new System.Windows.Forms.RichTextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.InsertButton = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.EvaluationBox = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // InputBox1
            // 
            this.InputBox1.Location = new System.Drawing.Point(62, 44);
            this.InputBox1.Multiline = false;
            this.InputBox1.Name = "InputBox1";
            this.InputBox1.Size = new System.Drawing.Size(395, 28);
            this.InputBox1.TabIndex = 0;
            this.InputBox1.Text = "";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(295, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "Select input ranges and sum the values !";
            // 
            // InputBox2
            // 
            this.InputBox2.Location = new System.Drawing.Point(62, 78);
            this.InputBox2.Multiline = false;
            this.InputBox2.Name = "InputBox2";
            this.InputBox2.Size = new System.Drawing.Size(395, 28);
            this.InputBox2.TabIndex = 2;
            this.InputBox2.Text = "";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Augend";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Addend";
            // 
            // DestinationBox
            // 
            this.DestinationBox.Location = new System.Drawing.Point(79, 148);
            this.DestinationBox.Multiline = false;
            this.DestinationBox.Name = "DestinationBox";
            this.DestinationBox.Size = new System.Drawing.Size(378, 28);
            this.DestinationBox.TabIndex = 5;
            this.DestinationBox.Text = "";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 151);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Destination";
            // 
            // InsertButton
            // 
            this.InsertButton.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.InsertButton.Location = new System.Drawing.Point(79, 194);
            this.InsertButton.Name = "InsertButton";
            this.InsertButton.Size = new System.Drawing.Size(219, 30);
            this.InsertButton.TabIndex = 7;
            this.InsertButton.Text = "Insert";
            this.InsertButton.UseVisualStyleBackColor = false;
            this.InsertButton.Click += new System.EventHandler(this.InsertButton_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 240);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(57, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Evaluation";
            // 
            // EvaluationBox
            // 
            this.EvaluationBox.Location = new System.Drawing.Point(79, 237);
            this.EvaluationBox.Name = "EvaluationBox";
            this.EvaluationBox.ReadOnly = true;
            this.EvaluationBox.Size = new System.Drawing.Size(378, 28);
            this.EvaluationBox.TabIndex = 9;
            this.EvaluationBox.Text = "";
            // 
            // RefEditForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(468, 278);
            this.Controls.Add(this.EvaluationBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.InsertButton);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.DestinationBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.InputBox2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.InputBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(474, 302);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(474, 302);
            this.Name = "RefEditForm";
            this.Text = "RefEdit";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox InputBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox InputBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox DestinationBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button InsertButton;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RichTextBox EvaluationBox;
    }
}