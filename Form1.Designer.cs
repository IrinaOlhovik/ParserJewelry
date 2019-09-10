namespace ParserApp
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
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.richTextBoxIdProducts = new System.Windows.Forms.RichTextBox();
            this.richTextBoxId = new System.Windows.Forms.RichTextBox();
            this.naturalIds = new System.Windows.Forms.RichTextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.textBoxPathLinks = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(445, 35);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 7;
            this.button1.Text = "Открыть";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBoxPath
            // 
            this.textBoxPath.Location = new System.Drawing.Point(12, 37);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.Size = new System.Drawing.Size(428, 20);
            this.textBoxPath.TabIndex = 6;
            this.textBoxPath.Text = "C:\\Users\\Admin\\Desktop";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Путь к .xlsx ";
            // 
            // button2
            // 
            this.button2.Cursor = System.Windows.Forms.Cursors.Help;
            this.button2.Location = new System.Drawing.Point(712, 474);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 13;
            this.button2.Text = "Парсить";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Cursor = System.Windows.Forms.Cursors.No;
            this.label8.Location = new System.Drawing.Point(422, 436);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(0, 13);
            this.label8.TabIndex = 14;
            this.label8.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Номера:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(422, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Артикулы:";
            // 
            // richTextBoxIdProducts
            // 
            this.richTextBoxIdProducts.Location = new System.Drawing.Point(12, 114);
            this.richTextBoxIdProducts.Name = "richTextBoxIdProducts";
            this.richTextBoxIdProducts.Size = new System.Drawing.Size(350, 167);
            this.richTextBoxIdProducts.TabIndex = 11;
            this.richTextBoxIdProducts.Text = "";
            // 
            // richTextBoxId
            // 
            this.richTextBoxId.Location = new System.Drawing.Point(425, 114);
            this.richTextBoxId.Name = "richTextBoxId";
            this.richTextBoxId.Size = new System.Drawing.Size(350, 167);
            this.richTextBoxId.TabIndex = 12;
            this.richTextBoxId.Text = "";
            // 
            // naturalIds
            // 
            this.naturalIds.Location = new System.Drawing.Point(12, 313);
            this.naturalIds.Name = "naturalIds";
            this.naturalIds.Size = new System.Drawing.Size(350, 179);
            this.naturalIds.TabIndex = 18;
            this.naturalIds.Text = "";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(12, 297);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(78, 13);
            this.label7.TabIndex = 19;
            this.label7.Text = "Натуральный:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(422, 316);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(118, 13);
            this.label6.TabIndex = 22;
            this.label6.Text = "Путь к .xml(описание):";
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Location = new System.Drawing.Point(425, 358);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(134, 23);
            this.button3.TabIndex = 21;
            this.button3.Text = "Добавить описание";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBoxPathLinks
            // 
            this.textBoxPathLinks.Location = new System.Drawing.Point(425, 332);
            this.textBoxPathLinks.Name = "textBoxPathLinks";
            this.textBoxPathLinks.Size = new System.Drawing.Size(350, 20);
            this.textBoxPathLinks.TabIndex = 20;
            this.textBoxPathLinks.Text = "https://amoresilver.kz/yandex_market.xml?hash_tag=13d4fc6fa705dbed0d972cc62cea8e9" +
    "c&sales_notes=&product_ids=&group_ids=&label_ids=&exclude_fields=&html_descripti" +
    "on=0&yandex_cpa=&process_presence_sure=";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Cursor = System.Windows.Forms.Cursors.No;
            this.label4.Location = new System.Drawing.Point(647, 479);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 13);
            this.label4.TabIndex = 23;
            this.label4.Text = "Парсинг...";
            this.label4.Visible = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(425, 450);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(362, 23);
            this.progressBar1.TabIndex = 24;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(794, 506);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.textBoxPathLinks);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.naturalIds);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.richTextBoxId);
            this.Controls.Add(this.richTextBoxIdProducts);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBoxPath);
            this.Name = "Form1";
            this.Text = "Парсер";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox richTextBoxIdProducts;
        private System.Windows.Forms.RichTextBox richTextBoxId;
        private System.Windows.Forms.RichTextBox naturalIds;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox textBoxPathLinks;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }
}

