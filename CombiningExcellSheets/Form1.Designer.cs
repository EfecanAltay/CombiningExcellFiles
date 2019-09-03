namespace CombiningExcellSheets
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
            this.btn_loadExcells = new System.Windows.Forms.Button();
            this.btn_join = new System.Windows.Forms.Button();
            this.listbox_excellPaths = new System.Windows.Forms.ListBox();
            this.btn_select_out = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lstbox_outPath = new System.Windows.Forms.ListBox();
            this.lstBox_output = new System.Windows.Forms.ListBox();
            this.btn_show_totalRowSize = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.lbl_rowCount = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_loadExcells
            // 
            this.btn_loadExcells.Location = new System.Drawing.Point(12, 12);
            this.btn_loadExcells.Name = "btn_loadExcells";
            this.btn_loadExcells.Size = new System.Drawing.Size(92, 43);
            this.btn_loadExcells.TabIndex = 0;
            this.btn_loadExcells.Text = "Excell Yükle";
            this.btn_loadExcells.UseVisualStyleBackColor = true;
            this.btn_loadExcells.Click += new System.EventHandler(this.Btn_loadExcells_Click);
            // 
            // btn_join
            // 
            this.btn_join.Location = new System.Drawing.Point(8, 269);
            this.btn_join.Name = "btn_join";
            this.btn_join.Size = new System.Drawing.Size(325, 40);
            this.btn_join.TabIndex = 1;
            this.btn_join.Text = "Birleştir";
            this.btn_join.UseVisualStyleBackColor = true;
            this.btn_join.Click += new System.EventHandler(this.Btn_join_Click);
            // 
            // listbox_excellPaths
            // 
            this.listbox_excellPaths.FormattingEnabled = true;
            this.listbox_excellPaths.Location = new System.Drawing.Point(12, 73);
            this.listbox_excellPaths.Name = "listbox_excellPaths";
            this.listbox_excellPaths.Size = new System.Drawing.Size(254, 108);
            this.listbox_excellPaths.TabIndex = 2;
            this.listbox_excellPaths.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Listbox_excellPaths_KeyDown);
            // 
            // btn_select_out
            // 
            this.btn_select_out.Location = new System.Drawing.Point(236, 12);
            this.btn_select_out.Name = "btn_select_out";
            this.btn_select_out.Size = new System.Drawing.Size(90, 43);
            this.btn_select_out.TabIndex = 3;
            this.btn_select_out.Text = "Çıktının Konumunu seç";
            this.btn_select_out.UseVisualStyleBackColor = true;
            this.btn_select_out.Click += new System.EventHandler(this.Btn_select_out_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 195);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Out Path :";
            // 
            // lstbox_outPath
            // 
            this.lstbox_outPath.BackColor = System.Drawing.SystemColors.Menu;
            this.lstbox_outPath.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lstbox_outPath.FormattingEnabled = true;
            this.lstbox_outPath.HorizontalScrollbar = true;
            this.lstbox_outPath.Location = new System.Drawing.Point(12, 211);
            this.lstbox_outPath.Name = "lstbox_outPath";
            this.lstbox_outPath.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.lstbox_outPath.Size = new System.Drawing.Size(321, 52);
            this.lstbox_outPath.TabIndex = 8;
            // 
            // lstBox_output
            // 
            this.lstBox_output.BackColor = System.Drawing.SystemColors.Menu;
            this.lstBox_output.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lstBox_output.FormattingEnabled = true;
            this.lstBox_output.HorizontalScrollbar = true;
            this.lstBox_output.Location = new System.Drawing.Point(8, 315);
            this.lstBox_output.Name = "lstBox_output";
            this.lstBox_output.Size = new System.Drawing.Size(321, 52);
            this.lstBox_output.TabIndex = 9;
            // 
            // btn_show_totalRowSize
            // 
            this.btn_show_totalRowSize.Location = new System.Drawing.Point(277, 103);
            this.btn_show_totalRowSize.Name = "btn_show_totalRowSize";
            this.btn_show_totalRowSize.Size = new System.Drawing.Size(56, 78);
            this.btn_show_totalRowSize.TabIndex = 10;
            this.btn_show_totalRowSize.Text = "Satır Boyutu Hesapla";
            this.btn_show_totalRowSize.UseVisualStyleBackColor = true;
            this.btn_show_totalRowSize.Click += new System.EventHandler(this.Btn_show_totalRowSize_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(272, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Toplam Satır";
            // 
            // lbl_rowCount
            // 
            this.lbl_rowCount.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_rowCount.Location = new System.Drawing.Point(275, 78);
            this.lbl_rowCount.Name = "lbl_rowCount";
            this.lbl_rowCount.Size = new System.Drawing.Size(63, 22);
            this.lbl_rowCount.TabIndex = 12;
            this.lbl_rowCount.Text = "---";
            this.lbl_rowCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(111, 13);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(119, 42);
            this.button1.TabIndex = 13;
            this.button1.Text = "Excell Ekle";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(338, 404);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lbl_rowCount);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_show_totalRowSize);
            this.Controls.Add(this.lstBox_output);
            this.Controls.Add(this.lstbox_outPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_select_out);
            this.Controls.Add(this.listbox_excellPaths);
            this.Controls.Add(this.btn_join);
            this.Controls.Add(this.btn_loadExcells);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "Excell Birleştirme";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_loadExcells;
        private System.Windows.Forms.Button btn_join;
        private System.Windows.Forms.ListBox listbox_excellPaths;
        private System.Windows.Forms.Button btn_select_out;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstbox_outPath;
        private System.Windows.Forms.ListBox lstBox_output;
        private System.Windows.Forms.Button btn_show_totalRowSize;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lbl_rowCount;
        private System.Windows.Forms.Button button1;
    }
}

