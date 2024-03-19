namespace maslova_cs_pr23
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
            this.btWord = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btWord
            // 
            this.btWord.Location = new System.Drawing.Point(13, 12);
            this.btWord.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btWord.Name = "btWord";
            this.btWord.Size = new System.Drawing.Size(304, 114);
            this.btWord.TabIndex = 0;
            this.btWord.Text = "В Ворд";
            this.btWord.UseVisualStyleBackColor = true;
            this.btWord.Click += new System.EventHandler(this.btWord_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(744, 134);
            this.Controls.Add(this.btWord);
            this.Font = new System.Drawing.Font("Bahnschrift SemiCondensed", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btWord;
    }
}

