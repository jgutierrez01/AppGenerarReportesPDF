namespace PDFReportes
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.lblEmbarque = new System.Windows.Forms.Label();
            this.CmbNumEmbarque = new System.Windows.Forms.ComboBox();
            this.lblRutaGuardar = new System.Windows.Forms.Label();
            this.btnSelectPathSave = new System.Windows.Forms.Button();
            this.btnGenerar = new System.Windows.Forms.Button();
            this.lblLogoSteelgo = new System.Windows.Forms.Label();
            this.folderGuardar = new System.Windows.Forms.FolderBrowserDialog();
            this.panelLoading = new System.Windows.Forms.Panel();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.backgroundWorkerGenerarPDF = new System.ComponentModel.BackgroundWorker();
            this.panelLoading.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // lblEmbarque
            // 
            this.lblEmbarque.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEmbarque.Location = new System.Drawing.Point(3, 80);
            this.lblEmbarque.Name = "lblEmbarque";
            this.lblEmbarque.Size = new System.Drawing.Size(176, 30);
            this.lblEmbarque.TabIndex = 2;
            this.lblEmbarque.Text = "Número de Embarque";
            // 
            // CmbNumEmbarque
            // 
            this.CmbNumEmbarque.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.CmbNumEmbarque.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.CmbNumEmbarque.DisplayMember = "NumeroEmbarque";
            this.CmbNumEmbarque.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CmbNumEmbarque.FormattingEnabled = true;
            this.CmbNumEmbarque.ItemHeight = 18;
            this.CmbNumEmbarque.Location = new System.Drawing.Point(6, 113);
            this.CmbNumEmbarque.Name = "CmbNumEmbarque";
            this.CmbNumEmbarque.Size = new System.Drawing.Size(306, 26);
            this.CmbNumEmbarque.TabIndex = 3;
            this.CmbNumEmbarque.ValueMember = "EmbarqueID";
            this.CmbNumEmbarque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.mayusculaNumEmbarque);
            // 
            // lblRutaGuardar
            // 
            this.lblRutaGuardar.AutoSize = true;
            this.lblRutaGuardar.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRutaGuardar.Location = new System.Drawing.Point(12, 202);
            this.lblRutaGuardar.Name = "lblRutaGuardar";
            this.lblRutaGuardar.Size = new System.Drawing.Size(189, 20);
            this.lblRutaGuardar.TabIndex = 4;
            this.lblRutaGuardar.Text = "Ninguna carpeta seleccionada";
            this.lblRutaGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnSelectPathSave
            // 
            this.btnSelectPathSave.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelectPathSave.Location = new System.Drawing.Point(6, 172);
            this.btnSelectPathSave.Name = "btnSelectPathSave";
            this.btnSelectPathSave.Size = new System.Drawing.Size(131, 27);
            this.btnSelectPathSave.TabIndex = 5;
            this.btnSelectPathSave.Text = "Seleccione Carpeta";
            this.btnSelectPathSave.UseVisualStyleBackColor = true;
            this.btnSelectPathSave.Click += new System.EventHandler(this.btnSelectPathSave_Click);
            // 
            // btnGenerar
            // 
            this.btnGenerar.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGenerar.Location = new System.Drawing.Point(113, 283);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Size = new System.Drawing.Size(97, 30);
            this.btnGenerar.TabIndex = 6;
            this.btnGenerar.Text = "Generar";
            this.btnGenerar.UseVisualStyleBackColor = true;
            this.btnGenerar.Click += new System.EventHandler(this.ClicGenerarPDF);
            // 
            // lblLogoSteelgo
            // 
            this.lblLogoSteelgo.Image = ((System.Drawing.Image)(resources.GetObject("lblLogoSteelgo.Image")));
            this.lblLogoSteelgo.Location = new System.Drawing.Point(3, 1);
            this.lblLogoSteelgo.Name = "lblLogoSteelgo";
            this.lblLogoSteelgo.Size = new System.Drawing.Size(198, 59);
            this.lblLogoSteelgo.TabIndex = 7;
            // 
            // folderGuardar
            // 
            this.folderGuardar.Description = "Elija carpeta para guardar PDF";
            // 
            // panelLoading
            // 
            this.panelLoading.BackColor = System.Drawing.Color.Transparent;
            this.panelLoading.Controls.Add(this.pictureBox2);
            this.panelLoading.Location = new System.Drawing.Point(0, 63);
            this.panelLoading.Name = "panelLoading";
            this.panelLoading.Size = new System.Drawing.Size(330, 356);
            this.panelLoading.TabIndex = 35;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.InitialImage = ((System.Drawing.Image)(resources.GetObject("pictureBox2.InitialImage")));
            this.pictureBox2.Location = new System.Drawing.Point(128, 116);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(76, 73);
            this.pictureBox2.TabIndex = 0;
            this.pictureBox2.TabStop = false;
            // 
            // backgroundWorkerGenerarPDF
            // 
            this.backgroundWorkerGenerarPDF.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerGenerarPDF_DoWork);
            this.backgroundWorkerGenerarPDF.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorkerGenerarPDF_ProgressChanged);
            this.backgroundWorkerGenerarPDF.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerGenerarPDF_RunWorkerCompleted);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(329, 420);
            this.Controls.Add(this.panelLoading);
            this.Controls.Add(this.lblLogoSteelgo);
            this.Controls.Add(this.btnGenerar);
            this.Controls.Add(this.btnSelectPathSave);
            this.Controls.Add(this.lblRutaGuardar);
            this.Controls.Add(this.CmbNumEmbarque);
            this.Controls.Add(this.lblEmbarque);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PDF Reportes";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panelLoading.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lblEmbarque;
        private System.Windows.Forms.ComboBox CmbNumEmbarque;
        private System.Windows.Forms.Label lblRutaGuardar;
        private System.Windows.Forms.Button btnSelectPathSave;
        private System.Windows.Forms.Button btnGenerar;
        private System.Windows.Forms.Label lblLogoSteelgo;
        private System.Windows.Forms.FolderBrowserDialog folderGuardar;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Panel panelLoading;
        private System.ComponentModel.BackgroundWorker backgroundWorkerGenerarPDF;
    }
}

