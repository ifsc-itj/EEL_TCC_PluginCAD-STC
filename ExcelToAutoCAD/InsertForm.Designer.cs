namespace ExcelToAutoCAD
{
    partial class InsertForm
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
            this.rbCoordinates = new System.Windows.Forms.RadioButton();
            this.rbScreen = new System.Windows.Forms.RadioButton();
            this.gbCoordinates = new System.Windows.Forms.GroupBox();
            this.txtCoordY = new System.Windows.Forms.TextBox();
            this.txtCoordX = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btOk = new System.Windows.Forms.Button();
            this.lbInfo = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btFileSearch = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.gbCoordinates.SuspendLayout();
            this.SuspendLayout();
            // 
            // rbCoordinates
            // 
            this.rbCoordinates.AutoSize = true;
            this.rbCoordinates.Location = new System.Drawing.Point(25, 141);
            this.rbCoordinates.Name = "rbCoordinates";
            this.rbCoordinates.Size = new System.Drawing.Size(131, 17);
            this.rbCoordinates.TabIndex = 0;
            this.rbCoordinates.TabStop = true;
            this.rbCoordinates.Text = "Digite as coordenadas";
            this.rbCoordinates.UseVisualStyleBackColor = true;
            this.rbCoordinates.CheckedChanged += new System.EventHandler(this.rbCoordinates_CheckedChanged);
            // 
            // rbScreen
            // 
            this.rbScreen.AutoSize = true;
            this.rbScreen.Location = new System.Drawing.Point(184, 141);
            this.rbScreen.Name = "rbScreen";
            this.rbScreen.Size = new System.Drawing.Size(123, 17);
            this.rbScreen.TabIndex = 1;
            this.rbScreen.TabStop = true;
            this.rbScreen.Text = "Insirir o ponto na tela";
            this.rbScreen.UseVisualStyleBackColor = true;
            this.rbScreen.CheckedChanged += new System.EventHandler(this.rbScreen_CheckedChanged);
            // 
            // gbCoordinates
            // 
            this.gbCoordinates.Controls.Add(this.txtCoordY);
            this.gbCoordinates.Controls.Add(this.txtCoordX);
            this.gbCoordinates.Controls.Add(this.label2);
            this.gbCoordinates.Controls.Add(this.label1);
            this.gbCoordinates.Location = new System.Drawing.Point(25, 173);
            this.gbCoordinates.Name = "gbCoordinates";
            this.gbCoordinates.Size = new System.Drawing.Size(282, 97);
            this.gbCoordinates.TabIndex = 2;
            this.gbCoordinates.TabStop = false;
            this.gbCoordinates.Text = "Entre com as coordenadas";
            // 
            // txtCoordY
            // 
            this.txtCoordY.Location = new System.Drawing.Point(87, 61);
            this.txtCoordY.Name = "txtCoordY";
            this.txtCoordY.Size = new System.Drawing.Size(100, 20);
            this.txtCoordY.TabIndex = 3;
            // 
            // txtCoordX
            // 
            this.txtCoordX.Location = new System.Drawing.Point(87, 23);
            this.txtCoordX.Name = "txtCoordX";
            this.txtCoordX.Size = new System.Drawing.Size(100, 20);
            this.txtCoordX.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 64);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Coordenada Y:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Coordenada X:";
            // 
            // btOk
            // 
            this.btOk.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btOk.Location = new System.Drawing.Point(25, 302);
            this.btOk.Name = "btOk";
            this.btOk.Size = new System.Drawing.Size(75, 23);
            this.btOk.TabIndex = 3;
            this.btOk.Text = "Ok";
            this.btOk.UseVisualStyleBackColor = true;
            this.btOk.Click += new System.EventHandler(this.btOk_Click);
            // 
            // lbInfo
            // 
            this.lbInfo.AutoSize = true;
            this.lbInfo.Location = new System.Drawing.Point(22, 273);
            this.lbInfo.Name = "lbInfo";
            this.lbInfo.Size = new System.Drawing.Size(16, 13);
            this.lbInfo.TabIndex = 4;
            this.lbInfo.Text = "...";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(31, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Inserir arquivo:";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(112, 48);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(418, 20);
            this.txtFilePath.TabIndex = 8;
            // 
            // btFileSearch
            // 
            this.btFileSearch.Location = new System.Drawing.Point(536, 46);
            this.btFileSearch.Name = "btFileSearch";
            this.btFileSearch.Size = new System.Drawing.Size(75, 23);
            this.btFileSearch.TabIndex = 9;
            this.btFileSearch.Text = "Buscar";
            this.btFileSearch.UseVisualStyleBackColor = true;
            this.btFileSearch.Click += new System.EventHandler(this.btFileSearch_Click);
            // 
            // btCancel
            // 
            this.btCancel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btCancel.Location = new System.Drawing.Point(112, 302);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(75, 23);
            this.btCancel.TabIndex = 10;
            this.btCancel.Text = "Cancelar";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(21, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 24);
            this.label5.TabIndex = 11;
            this.label5.Text = "ExFill";
            // 
            // InsertForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(633, 347);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btFileSearch);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lbInfo);
            this.Controls.Add(this.btOk);
            this.Controls.Add(this.gbCoordinates);
            this.Controls.Add(this.rbScreen);
            this.Controls.Add(this.rbCoordinates);
            this.MaximizeBox = false;
            this.Name = "InsertForm";
            this.ShowIcon = false;
            this.Text = "ExFill";
            this.gbCoordinates.ResumeLayout(false);
            this.gbCoordinates.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton rbCoordinates;
        private System.Windows.Forms.RadioButton rbScreen;
        private System.Windows.Forms.GroupBox gbCoordinates;
        private System.Windows.Forms.Button btOk;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbInfo;
        private System.Windows.Forms.TextBox txtCoordY;
        private System.Windows.Forms.TextBox txtCoordX;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btFileSearch;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Label label5;
    }
}