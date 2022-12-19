namespace VeriTabanıV1._3
{
    partial class Form2ogisleri
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btn_kaydet = new System.Windows.Forms.Button();
            this.btn_sil = new System.Windows.Forms.Button();
            this.txt_id = new System.Windows.Forms.TextBox();
            this.txt_isim = new System.Windows.Forms.TextBox();
            this.txt_soyisim = new System.Windows.Forms.TextBox();
            this.txt_sinif = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(47, 166);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(670, 390);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // btn_kaydet
            // 
            this.btn_kaydet.Location = new System.Drawing.Point(757, 84);
            this.btn_kaydet.Name = "btn_kaydet";
            this.btn_kaydet.Size = new System.Drawing.Size(146, 36);
            this.btn_kaydet.TabIndex = 1;
            this.btn_kaydet.Text = "Kaydet";
            this.btn_kaydet.UseVisualStyleBackColor = true;
            this.btn_kaydet.Click += new System.EventHandler(this.buttonKayıt_Click);
            // 
            // btn_sil
            // 
            this.btn_sil.Location = new System.Drawing.Point(757, 138);
            this.btn_sil.Name = "btn_sil";
            this.btn_sil.Size = new System.Drawing.Size(146, 36);
            this.btn_sil.TabIndex = 2;
            this.btn_sil.Text = "Sil";
            this.btn_sil.UseVisualStyleBackColor = true;
            // 
            // txt_id
            // 
            this.txt_id.Location = new System.Drawing.Point(779, 293);
            this.txt_id.Name = "txt_id";
            this.txt_id.Size = new System.Drawing.Size(100, 22);
            this.txt_id.TabIndex = 3;
            // 
            // txt_isim
            // 
            this.txt_isim.Location = new System.Drawing.Point(779, 321);
            this.txt_isim.Name = "txt_isim";
            this.txt_isim.Size = new System.Drawing.Size(100, 22);
            this.txt_isim.TabIndex = 4;
            // 
            // txt_soyisim
            // 
            this.txt_soyisim.Location = new System.Drawing.Point(779, 349);
            this.txt_soyisim.Name = "txt_soyisim";
            this.txt_soyisim.Size = new System.Drawing.Size(100, 22);
            this.txt_soyisim.TabIndex = 5;
            // 
            // txt_sinif
            // 
            this.txt_sinif.Location = new System.Drawing.Point(779, 377);
            this.txt_sinif.Name = "txt_sinif";
            this.txt_sinif.Size = new System.Drawing.Size(100, 22);
            this.txt_sinif.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(696, 101);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 16);
            this.label1.TabIndex = 7;
            this.label1.Text = "label1";
            // 
            // Form2ogisleri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 549);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txt_sinif);
            this.Controls.Add(this.txt_soyisim);
            this.Controls.Add(this.txt_isim);
            this.Controls.Add(this.txt_id);
            this.Controls.Add(this.btn_sil);
            this.Controls.Add(this.btn_kaydet);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form2ogisleri";
            this.Text = "Form2ogisleri";
            this.Load += new System.EventHandler(this.Form2ogisleri_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btn_kaydet;
        private System.Windows.Forms.Button btn_sil;
        private System.Windows.Forms.TextBox txt_id;
        private System.Windows.Forms.TextBox txt_isim;
        private System.Windows.Forms.TextBox txt_soyisim;
        private System.Windows.Forms.TextBox txt_sinif;
        private System.Windows.Forms.Label label1;
    }
}