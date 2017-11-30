namespace Multiple_Excel_Merging_Tool
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
            this.button_browse1 = new System.Windows.Forms.Button();
            this.SelectedExcelFiles = new System.Windows.Forms.ListBox();
            this.button_merge = new System.Windows.Forms.Button();
            this.button_reset = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button_browse1
            // 
            this.button_browse1.Location = new System.Drawing.Point(38, 37);
            this.button_browse1.Name = "button_browse1";
            this.button_browse1.Size = new System.Drawing.Size(97, 31);
            this.button_browse1.TabIndex = 0;
            this.button_browse1.Text = "Browse Excel 1";
            this.button_browse1.UseVisualStyleBackColor = true;
            this.button_browse1.Click += new System.EventHandler(this.button_browse1_Click);
            // 
            // SelectedExcelFiles
            // 
            this.SelectedExcelFiles.FormattingEnabled = true;
            this.SelectedExcelFiles.Location = new System.Drawing.Point(38, 75);
            this.SelectedExcelFiles.Name = "SelectedExcelFiles";
            this.SelectedExcelFiles.Size = new System.Drawing.Size(628, 108);
            this.SelectedExcelFiles.TabIndex = 4;
            // 
            // button_merge
            // 
            this.button_merge.Location = new System.Drawing.Point(38, 259);
            this.button_merge.Name = "button_merge";
            this.button_merge.Size = new System.Drawing.Size(164, 31);
            this.button_merge.TabIndex = 5;
            this.button_merge.Text = "Merge Multiple Excel Files";
            this.button_merge.UseVisualStyleBackColor = true;
            this.button_merge.Click += new System.EventHandler(this.button_merge_Click);
            // 
            // button_reset
            // 
            this.button_reset.Location = new System.Drawing.Point(569, 259);
            this.button_reset.Name = "button_reset";
            this.button_reset.Size = new System.Drawing.Size(97, 31);
            this.button_reset.TabIndex = 6;
            this.button_reset.Text = "Reset";
            this.button_reset.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(697, 328);
            this.Controls.Add(this.button_reset);
            this.Controls.Add(this.button_merge);
            this.Controls.Add(this.SelectedExcelFiles);
            this.Controls.Add(this.button_browse1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button_browse1;
        private System.Windows.Forms.ListBox SelectedExcelFiles;
        private System.Windows.Forms.Button button_merge;
        private System.Windows.Forms.Button button_reset;
    }
}

