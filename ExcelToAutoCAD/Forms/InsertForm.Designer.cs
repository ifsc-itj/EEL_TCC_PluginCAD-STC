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
            this.helpProvider1 = new System.Windows.Forms.HelpProvider();
            this.TBBus = new System.Windows.Forms.TextBox();
            this.CBGrounding = new System.Windows.Forms.ComboBox();
            this.TBSource = new System.Windows.Forms.TextBox();
            this.CBTube = new System.Windows.Forms.ComboBox();
            this.CBVoltage = new System.Windows.Forms.ComboBox();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageMain = new System.Windows.Forms.TabPage();
            this.label27 = new System.Windows.Forms.Label();
            this.cbSheets = new System.Windows.Forms.ComboBox();
            this.gbCoordinates = new System.Windows.Forms.GroupBox();
            this.txtCoordY = new System.Windows.Forms.TextBox();
            this.txtCoordX = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btCancel = new System.Windows.Forms.Button();
            this.lbInfo = new System.Windows.Forms.Label();
            this.rbScreen = new System.Windows.Forms.RadioButton();
            this.rbCoordinates = new System.Windows.Forms.RadioButton();
            this.label5 = new System.Windows.Forms.Label();
            this.btFileSearch = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btOk = new System.Windows.Forms.Button();
            this.tabPageConfig = new System.Windows.Forms.TabPage();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.tbPhaseC = new System.Windows.Forms.TextBox();
            this.tbPhaseB = new System.Windows.Forms.TextBox();
            this.tbPhaseA = new System.Windows.Forms.TextBox();
            this.label36 = new System.Windows.Forms.Label();
            this.label35 = new System.Windows.Forms.Label();
            this.label34 = new System.Windows.Forms.Label();
            this.btCancelConfig = new System.Windows.Forms.Button();
            this.btApplyConfig = new System.Windows.Forms.Button();
            this.label18 = new System.Windows.Forms.Label();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.CBFontName = new System.Windows.Forms.ComboBox();
            this.TBFontSize = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.TBScale = new System.Windows.Forms.TextBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.label22 = new System.Windows.Forms.Label();
            this.TBGroundLayer = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.TBWiringLayer = new System.Windows.Forms.TextBox();
            this.TBNeutralLayer = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.TBCircuitLayer = new System.Windows.Forms.TextBox();
            this.TBBusLayer = new System.Windows.Forms.TextBox();
            this.label26 = new System.Windows.Forms.Label();
            this.tabPageMaterialList = new System.Windows.Forms.TabPage();
            this.btCADExport = new System.Windows.Forms.Button();
            this.btExcelExport = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.lbNotes = new System.Windows.Forms.Label();
            this.dgvMaterialList = new System.Windows.Forms.DataGridView();
            this.DescriptionList = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.QuantityList = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnUnit = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ManufacturerList = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CodeList = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ObsList = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btMatList = new System.Windows.Forms.Button();
            this.labelArqPath = new System.Windows.Forms.Label();
            this.lbArquivoSelect = new System.Windows.Forms.Label();
            this.tabPageOpMaterial = new System.Windows.Forms.TabPage();
            this.tbRestart = new System.Windows.Forms.Button();
            this.btSave = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.tbSPDCode = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.tbTerminalCode = new System.Windows.Forms.TextBox();
            this.tbRCDCode = new System.Windows.Forms.TextBox();
            this.tbBreakerCode = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tbSPDManuf = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.tbTerminalManuf = new System.Windows.Forms.TextBox();
            this.tbRCDManuf = new System.Windows.Forms.TextBox();
            this.tbBreakerManuf = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.tabPageLoads = new System.Windows.Forms.TabPage();
            this.btLoadFileFromExcel = new System.Windows.Forms.Button();
            this.btLoadToCad = new System.Windows.Forms.Button();
            this.dgvLoads = new System.Windows.Forms.DataGridView();
            this.tabPagePanelMount = new System.Windows.Forms.TabPage();
            this.groupBox12 = new System.Windows.Forms.GroupBox();
            this.groupBox13 = new System.Windows.Forms.GroupBox();
            this.label37 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.btMountPanel = new System.Windows.Forms.Button();
            this.tbAbout = new System.Windows.Forms.TabPage();
            this.btManual = new System.Windows.Forms.Button();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.linkLabel3 = new System.Windows.Forms.LinkLabel();
            this.linkLabel4 = new System.Windows.Forms.LinkLabel();
            this.lbVersion = new System.Windows.Forms.Label();
            this.label32 = new System.Windows.Forms.Label();
            this.label31 = new System.Windows.Forms.Label();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.label30 = new System.Windows.Forms.Label();
            this.label29 = new System.Windows.Forms.Label();
            this.label28 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPageMain.SuspendLayout();
            this.gbCoordinates.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPageConfig.SuspendLayout();
            this.groupBox11.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.tabPageMaterialList.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMaterialList)).BeginInit();
            this.tabPageOpMaterial.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabPageLoads.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLoads)).BeginInit();
            this.tabPagePanelMount.SuspendLayout();
            this.groupBox12.SuspendLayout();
            this.tbAbout.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.SuspendLayout();
            // 
            // TBBus
            // 
            this.helpProvider1.SetHelpString(this.TBBus, "Descreva de onde vem a alimentação do quadro");
            this.TBBus.Location = new System.Drawing.Point(79, 127);
            this.TBBus.Name = "TBBus";
            this.helpProvider1.SetShowHelp(this.TBBus, true);
            this.TBBus.Size = new System.Drawing.Size(420, 20);
            this.TBBus.TabIndex = 21;
            // 
            // CBGrounding
            // 
            this.CBGrounding.FormattingEnabled = true;
            this.helpProvider1.SetHelpString(this.CBGrounding, "Defina o sistema de aterramento, segundo NBR 5410");
            this.CBGrounding.Items.AddRange(new object[] {
            "Sistema de aterramento TN NBR-5410",
            "Sistema de aterramento TN-S NBR-5410",
            "Sistema de aterramento TN-C NBR-5410",
            "Sistema de aterramento TN-C-S NBR-5410",
            "Sistema de aterramento TT NBR-5410",
            "Sistema de aterramento IT NBR-5410"});
            this.CBGrounding.Location = new System.Drawing.Point(79, 100);
            this.CBGrounding.Name = "CBGrounding";
            this.helpProvider1.SetShowHelp(this.CBGrounding, true);
            this.CBGrounding.Size = new System.Drawing.Size(420, 21);
            this.CBGrounding.TabIndex = 19;
            // 
            // TBSource
            // 
            this.helpProvider1.SetHelpString(this.TBSource, "Descreva de onde vem a alimentação do quadro");
            this.TBSource.Location = new System.Drawing.Point(79, 74);
            this.TBSource.Name = "TBSource";
            this.helpProvider1.SetShowHelp(this.TBSource, true);
            this.TBSource.Size = new System.Drawing.Size(420, 20);
            this.TBSource.TabIndex = 16;
            // 
            // CBTube
            // 
            this.CBTube.FormattingEnabled = true;
            this.helpProvider1.SetHelpString(this.CBTube, "Insira o eletroduto de alimentação para o quadro");
            this.CBTube.Items.AddRange(new object[] {
            "1/2\"",
            "3/4\"",
            "1\"",
            "1.1/4\"",
            "1.1/2\"",
            "2\"",
            "2.1/2\"",
            "3\"",
            "4\""});
            this.CBTube.Location = new System.Drawing.Point(79, 47);
            this.CBTube.Name = "CBTube";
            this.helpProvider1.SetShowHelp(this.CBTube, true);
            this.CBTube.Size = new System.Drawing.Size(181, 21);
            this.CBTube.TabIndex = 17;
            // 
            // CBVoltage
            // 
            this.CBVoltage.FormattingEnabled = true;
            this.helpProvider1.SetHelpString(this.CBVoltage, "Defina a tensão entre fase-fase");
            this.CBVoltage.Items.AddRange(new object[] {
            "110 V",
            "127 V",
            "220 V",
            "240 V",
            "380 V",
            "400 V"});
            this.CBVoltage.Location = new System.Drawing.Point(79, 20);
            this.CBVoltage.Name = "CBVoltage";
            this.helpProvider1.SetShowHelp(this.CBVoltage, true);
            this.CBVoltage.Size = new System.Drawing.Size(181, 21);
            this.CBVoltage.TabIndex = 12;
            // 
            // txtFilePath
            // 
            this.txtFilePath.BackColor = System.Drawing.SystemColors.Window;
            this.helpProvider1.SetHelpString(this.txtFilePath, "Selecione o arquivo padrão .xlsx");
            this.txtFilePath.Location = new System.Drawing.Point(123, 24);
            this.txtFilePath.Name = "txtFilePath";
            this.helpProvider1.SetShowHelp(this.txtFilePath, true);
            this.txtFilePath.Size = new System.Drawing.Size(418, 20);
            this.txtFilePath.TabIndex = 21;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageMain);
            this.tabControl1.Controls.Add(this.tabPageConfig);
            this.tabControl1.Controls.Add(this.tabPageMaterialList);
            this.tabControl1.Controls.Add(this.tabPageOpMaterial);
            this.tabControl1.Controls.Add(this.tabPageLoads);
            this.tabControl1.Controls.Add(this.tabPagePanelMount);
            this.tabControl1.Controls.Add(this.tbAbout);
            this.tabControl1.HotTrack = true;
            this.tabControl1.Location = new System.Drawing.Point(3, 4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(705, 481);
            this.tabControl1.TabIndex = 15;
            // 
            // tabPageMain
            // 
            this.tabPageMain.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageMain.Controls.Add(this.label27);
            this.tabPageMain.Controls.Add(this.cbSheets);
            this.tabPageMain.Controls.Add(this.gbCoordinates);
            this.tabPageMain.Controls.Add(this.groupBox1);
            this.tabPageMain.Controls.Add(this.btCancel);
            this.tabPageMain.Controls.Add(this.lbInfo);
            this.tabPageMain.Controls.Add(this.rbScreen);
            this.tabPageMain.Controls.Add(this.rbCoordinates);
            this.tabPageMain.Controls.Add(this.label5);
            this.tabPageMain.Controls.Add(this.btFileSearch);
            this.tabPageMain.Controls.Add(this.txtFilePath);
            this.tabPageMain.Controls.Add(this.label4);
            this.tabPageMain.Controls.Add(this.btOk);
            this.tabPageMain.Location = new System.Drawing.Point(4, 22);
            this.tabPageMain.Name = "tabPageMain";
            this.tabPageMain.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMain.Size = new System.Drawing.Size(697, 455);
            this.tabPageMain.TabIndex = 0;
            this.tabPageMain.Text = "Geral";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(12, 53);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(105, 13);
            this.label27.TabIndex = 27;
            this.label27.Text = "Selecione a planilha;";
            // 
            // cbSheets
            // 
            this.cbSheets.FormattingEnabled = true;
            this.cbSheets.Location = new System.Drawing.Point(123, 50);
            this.cbSheets.Name = "cbSheets";
            this.cbSheets.Size = new System.Drawing.Size(121, 21);
            this.cbSheets.TabIndex = 26;
            this.cbSheets.SelectedIndexChanged += new System.EventHandler(this.cbSheets_SelectedIndexChanged);
            // 
            // gbCoordinates
            // 
            this.gbCoordinates.Controls.Add(this.txtCoordY);
            this.gbCoordinates.Controls.Add(this.txtCoordX);
            this.gbCoordinates.Controls.Add(this.label2);
            this.gbCoordinates.Controls.Add(this.label1);
            this.gbCoordinates.Location = new System.Drawing.Point(36, 304);
            this.gbCoordinates.Name = "gbCoordinates";
            this.gbCoordinates.Size = new System.Drawing.Size(205, 94);
            this.gbCoordinates.TabIndex = 17;
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
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.TBBus);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.CBGrounding);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.CBTube);
            this.groupBox1.Controls.Add(this.TBSource);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.CBVoltage);
            this.groupBox1.Location = new System.Drawing.Point(36, 99);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(505, 176);
            this.groupBox1.TabIndex = 25;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Configurações gerais";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(7, 130);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(64, 13);
            this.label9.TabIndex = 20;
            this.label9.Text = "Barramento:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 103);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(67, 13);
            this.label8.TabIndex = 18;
            this.label8.Text = "Aterramento:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(58, 13);
            this.label7.TabIndex = 15;
            this.label7.Text = "Eletroduto:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(7, 77);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Alimentação:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Tensão F-F:";
            // 
            // btCancel
            // 
            this.btCancel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btCancel.Location = new System.Drawing.Point(123, 417);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(75, 23);
            this.btCancel.TabIndex = 23;
            this.btCancel.Text = "Cancelar";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click_1);
            // 
            // lbInfo
            // 
            this.lbInfo.AutoSize = true;
            this.lbInfo.Location = new System.Drawing.Point(42, 401);
            this.lbInfo.Name = "lbInfo";
            this.lbInfo.Size = new System.Drawing.Size(16, 13);
            this.lbInfo.TabIndex = 19;
            this.lbInfo.Text = "...";
            // 
            // rbScreen
            // 
            this.rbScreen.AutoSize = true;
            this.rbScreen.Checked = true;
            this.rbScreen.Location = new System.Drawing.Point(173, 281);
            this.rbScreen.Name = "rbScreen";
            this.rbScreen.Size = new System.Drawing.Size(123, 17);
            this.rbScreen.TabIndex = 16;
            this.rbScreen.TabStop = true;
            this.rbScreen.Text = "Insirir o ponto na tela";
            this.rbScreen.UseVisualStyleBackColor = true;
            this.rbScreen.CheckedChanged += new System.EventHandler(this.rbScreen_CheckedChanged_1);
            // 
            // rbCoordinates
            // 
            this.rbCoordinates.AutoSize = true;
            this.rbCoordinates.Location = new System.Drawing.Point(36, 281);
            this.rbCoordinates.Name = "rbCoordinates";
            this.rbCoordinates.Size = new System.Drawing.Size(131, 17);
            this.rbCoordinates.TabIndex = 15;
            this.rbCoordinates.Text = "Digite as coordenadas";
            this.rbCoordinates.UseVisualStyleBackColor = true;
            this.rbCoordinates.CheckedChanged += new System.EventHandler(this.rbCoordinates_CheckedChanged_1);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(527, 416);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(164, 24);
            this.label5.TabIndex = 24;
            this.label5.Text = "Sheet To CAD ®️";
            // 
            // btFileSearch
            // 
            this.btFileSearch.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btFileSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.btFileSearch.Location = new System.Drawing.Point(547, 22);
            this.btFileSearch.Name = "btFileSearch";
            this.btFileSearch.Size = new System.Drawing.Size(75, 23);
            this.btFileSearch.TabIndex = 22;
            this.btFileSearch.Text = "Buscar";
            this.btFileSearch.UseVisualStyleBackColor = false;
            this.btFileSearch.Click += new System.EventHandler(this.btFileSearch_Click_1);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(41, 24);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "Inserir arquivo:";
            // 
            // btOk
            // 
            this.btOk.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btOk.Location = new System.Drawing.Point(36, 417);
            this.btOk.Name = "btOk";
            this.btOk.Size = new System.Drawing.Size(75, 23);
            this.btOk.TabIndex = 18;
            this.btOk.Text = "Ok";
            this.btOk.UseVisualStyleBackColor = true;
            this.btOk.Click += new System.EventHandler(this.btOk_Click_1);
            // 
            // tabPageConfig
            // 
            this.tabPageConfig.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageConfig.Controls.Add(this.groupBox11);
            this.tabPageConfig.Controls.Add(this.btCancelConfig);
            this.tabPageConfig.Controls.Add(this.btApplyConfig);
            this.tabPageConfig.Controls.Add(this.label18);
            this.tabPageConfig.Controls.Add(this.groupBox6);
            this.tabPageConfig.Controls.Add(this.TBScale);
            this.tabPageConfig.Controls.Add(this.groupBox7);
            this.tabPageConfig.Location = new System.Drawing.Point(4, 22);
            this.tabPageConfig.Name = "tabPageConfig";
            this.tabPageConfig.Size = new System.Drawing.Size(697, 455);
            this.tabPageConfig.TabIndex = 2;
            this.tabPageConfig.Text = "Configurações";
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.tbPhaseC);
            this.groupBox11.Controls.Add(this.tbPhaseB);
            this.groupBox11.Controls.Add(this.tbPhaseA);
            this.groupBox11.Controls.Add(this.label36);
            this.groupBox11.Controls.Add(this.label35);
            this.groupBox11.Controls.Add(this.label34);
            this.groupBox11.Location = new System.Drawing.Point(15, 246);
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.Size = new System.Drawing.Size(175, 113);
            this.groupBox11.TabIndex = 28;
            this.groupBox11.TabStop = false;
            this.groupBox11.Text = "Fases";
            // 
            // tbPhaseC
            // 
            this.tbPhaseC.Location = new System.Drawing.Point(55, 81);
            this.tbPhaseC.Name = "tbPhaseC";
            this.tbPhaseC.Size = new System.Drawing.Size(100, 20);
            this.tbPhaseC.TabIndex = 5;
            // 
            // tbPhaseB
            // 
            this.tbPhaseB.Location = new System.Drawing.Point(55, 53);
            this.tbPhaseB.Name = "tbPhaseB";
            this.tbPhaseB.Size = new System.Drawing.Size(100, 20);
            this.tbPhaseB.TabIndex = 4;
            // 
            // tbPhaseA
            // 
            this.tbPhaseA.Location = new System.Drawing.Point(56, 22);
            this.tbPhaseA.Name = "tbPhaseA";
            this.tbPhaseA.Size = new System.Drawing.Size(100, 20);
            this.tbPhaseA.TabIndex = 3;
            // 
            // label36
            // 
            this.label36.AutoSize = true;
            this.label36.Location = new System.Drawing.Point(7, 81);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(43, 13);
            this.label36.TabIndex = 2;
            this.label36.Text = "Fase T:";
            // 
            // label35
            // 
            this.label35.AutoSize = true;
            this.label35.Location = new System.Drawing.Point(7, 53);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(43, 13);
            this.label35.TabIndex = 1;
            this.label35.Text = "Fase S:";
            // 
            // label34
            // 
            this.label34.AutoSize = true;
            this.label34.Location = new System.Drawing.Point(7, 22);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(44, 13);
            this.label34.TabIndex = 0;
            this.label34.Text = "Fase R:";
            // 
            // btCancelConfig
            // 
            this.btCancelConfig.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btCancelConfig.Location = new System.Drawing.Point(610, 419);
            this.btCancelConfig.Name = "btCancelConfig";
            this.btCancelConfig.Size = new System.Drawing.Size(75, 23);
            this.btCancelConfig.TabIndex = 27;
            this.btCancelConfig.Text = "Cancelar";
            this.btCancelConfig.UseVisualStyleBackColor = true;
            this.btCancelConfig.Click += new System.EventHandler(this.btCancelConfig_Click);
            // 
            // btApplyConfig
            // 
            this.btApplyConfig.Location = new System.Drawing.Point(15, 419);
            this.btApplyConfig.Name = "btApplyConfig";
            this.btApplyConfig.Size = new System.Drawing.Size(75, 23);
            this.btApplyConfig.TabIndex = 26;
            this.btApplyConfig.Text = "Salvar";
            this.btApplyConfig.UseVisualStyleBackColor = true;
            this.btApplyConfig.Click += new System.EventHandler(this.btApplyConfig_Click);
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(27, 15);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(42, 13);
            this.label18.TabIndex = 22;
            this.label18.Text = "Escala:";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.CBFontName);
            this.groupBox6.Controls.Add(this.TBFontSize);
            this.groupBox6.Controls.Add(this.label20);
            this.groupBox6.Controls.Add(this.label21);
            this.groupBox6.Location = new System.Drawing.Point(357, 38);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(269, 94);
            this.groupBox6.TabIndex = 24;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Textos";
            // 
            // CBFontName
            // 
            this.CBFontName.FormattingEnabled = true;
            this.CBFontName.Location = new System.Drawing.Point(109, 28);
            this.CBFontName.Name = "CBFontName";
            this.CBFontName.Size = new System.Drawing.Size(143, 21);
            this.CBFontName.TabIndex = 18;
            // 
            // TBFontSize
            // 
            this.TBFontSize.Location = new System.Drawing.Point(109, 55);
            this.TBFontSize.Name = "TBFontSize";
            this.TBFontSize.Size = new System.Drawing.Size(143, 20);
            this.TBFontSize.TabIndex = 16;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(9, 58);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(97, 13);
            this.label20.TabIndex = 17;
            this.label20.Text = "Tamanho da fonte:";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(9, 32);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(72, 13);
            this.label21.TabIndex = 5;
            this.label21.Text = "Tipo de texto:";
            // 
            // TBScale
            // 
            this.TBScale.Location = new System.Drawing.Point(70, 12);
            this.TBScale.Name = "TBScale";
            this.TBScale.Size = new System.Drawing.Size(36, 20);
            this.TBScale.TabIndex = 21;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.label22);
            this.groupBox7.Controls.Add(this.TBGroundLayer);
            this.groupBox7.Controls.Add(this.label23);
            this.groupBox7.Controls.Add(this.label24);
            this.groupBox7.Controls.Add(this.TBWiringLayer);
            this.groupBox7.Controls.Add(this.TBNeutralLayer);
            this.groupBox7.Controls.Add(this.label25);
            this.groupBox7.Controls.Add(this.TBCircuitLayer);
            this.groupBox7.Controls.Add(this.TBBusLayer);
            this.groupBox7.Controls.Add(this.label26);
            this.groupBox7.Location = new System.Drawing.Point(15, 38);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(336, 202);
            this.groupBox7.TabIndex = 23;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Layers";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label22.Location = new System.Drawing.Point(12, 133);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(64, 13);
            this.label22.TabIndex = 15;
            this.label22.Text = "Layer Terra:";
            // 
            // TBGroundLayer
            // 
            this.TBGroundLayer.Location = new System.Drawing.Point(111, 133);
            this.TBGroundLayer.Name = "TBGroundLayer";
            this.TBGroundLayer.Size = new System.Drawing.Size(209, 20);
            this.TBGroundLayer.TabIndex = 14;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(12, 159);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(71, 13);
            this.label23.TabIndex = 7;
            this.label23.Text = "Layer Fiação:";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label24.Location = new System.Drawing.Point(12, 107);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(71, 13);
            this.label24.TabIndex = 13;
            this.label24.Text = "Layer Neutro:";
            // 
            // TBWiringLayer
            // 
            this.TBWiringLayer.Location = new System.Drawing.Point(111, 159);
            this.TBWiringLayer.Name = "TBWiringLayer";
            this.TBWiringLayer.Size = new System.Drawing.Size(209, 20);
            this.TBWiringLayer.TabIndex = 6;
            // 
            // TBNeutralLayer
            // 
            this.TBNeutralLayer.Location = new System.Drawing.Point(111, 107);
            this.TBNeutralLayer.Name = "TBNeutralLayer";
            this.TBNeutralLayer.Size = new System.Drawing.Size(209, 20);
            this.TBNeutralLayer.TabIndex = 12;
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label25.Location = new System.Drawing.Point(12, 81);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(93, 13);
            this.label25.TabIndex = 11;
            this.label25.Text = "Layer Barramento:";
            // 
            // TBCircuitLayer
            // 
            this.TBCircuitLayer.Location = new System.Drawing.Point(111, 55);
            this.TBCircuitLayer.Name = "TBCircuitLayer";
            this.TBCircuitLayer.Size = new System.Drawing.Size(209, 20);
            this.TBCircuitLayer.TabIndex = 4;
            // 
            // TBBusLayer
            // 
            this.TBBusLayer.Location = new System.Drawing.Point(111, 81);
            this.TBBusLayer.Name = "TBBusLayer";
            this.TBBusLayer.Size = new System.Drawing.Size(209, 20);
            this.TBBusLayer.TabIndex = 10;
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(12, 55);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(79, 13);
            this.label26.TabIndex = 5;
            this.label26.Text = "Layer Circuitos:";
            // 
            // tabPageMaterialList
            // 
            this.tabPageMaterialList.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageMaterialList.Controls.Add(this.btCADExport);
            this.tabPageMaterialList.Controls.Add(this.btExcelExport);
            this.tabPageMaterialList.Controls.Add(this.groupBox4);
            this.tabPageMaterialList.Controls.Add(this.dgvMaterialList);
            this.tabPageMaterialList.Controls.Add(this.btMatList);
            this.tabPageMaterialList.Controls.Add(this.labelArqPath);
            this.tabPageMaterialList.Controls.Add(this.lbArquivoSelect);
            this.tabPageMaterialList.Location = new System.Drawing.Point(4, 22);
            this.tabPageMaterialList.Name = "tabPageMaterialList";
            this.tabPageMaterialList.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMaterialList.Size = new System.Drawing.Size(697, 455);
            this.tabPageMaterialList.TabIndex = 1;
            this.tabPageMaterialList.Text = "Lista de material";
            // 
            // btCADExport
            // 
            this.btCADExport.Location = new System.Drawing.Point(196, 409);
            this.btCADExport.Name = "btCADExport";
            this.btCADExport.Size = new System.Drawing.Size(75, 23);
            this.btCADExport.TabIndex = 6;
            this.btCADExport.Text = "CAD";
            this.btCADExport.UseVisualStyleBackColor = true;
            this.btCADExport.Click += new System.EventHandler(this.btCADExport_Click);
            // 
            // btExcelExport
            // 
            this.btExcelExport.Location = new System.Drawing.Point(115, 409);
            this.btExcelExport.Name = "btExcelExport";
            this.btExcelExport.Size = new System.Drawing.Size(75, 23);
            this.btExcelExport.TabIndex = 5;
            this.btExcelExport.Text = "Excel";
            this.btExcelExport.UseVisualStyleBackColor = true;
            this.btExcelExport.Click += new System.EventHandler(this.btExcelExport_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.lbNotes);
            this.groupBox4.Location = new System.Drawing.Point(378, 397);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(265, 55);
            this.groupBox4.TabIndex = 4;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Observações";
            // 
            // lbNotes
            // 
            this.lbNotes.AutoSize = true;
            this.lbNotes.Location = new System.Drawing.Point(7, 21);
            this.lbNotes.Name = "lbNotes";
            this.lbNotes.Size = new System.Drawing.Size(16, 13);
            this.lbNotes.TabIndex = 0;
            this.lbNotes.Text = "...";
            // 
            // dgvMaterialList
            // 
            this.dgvMaterialList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMaterialList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DescriptionList,
            this.QuantityList,
            this.ColumnUnit,
            this.ManufacturerList,
            this.CodeList,
            this.ObsList});
            this.dgvMaterialList.Location = new System.Drawing.Point(18, 60);
            this.dgvMaterialList.Name = "dgvMaterialList";
            this.dgvMaterialList.Size = new System.Drawing.Size(673, 331);
            this.dgvMaterialList.TabIndex = 3;
            // 
            // DescriptionList
            // 
            this.DescriptionList.HeaderText = "Descrição";
            this.DescriptionList.MinimumWidth = 15;
            this.DescriptionList.Name = "DescriptionList";
            this.DescriptionList.Width = 180;
            // 
            // QuantityList
            // 
            this.QuantityList.HeaderText = "Quantidade";
            this.QuantityList.Name = "QuantityList";
            // 
            // ColumnUnit
            // 
            this.ColumnUnit.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.ColumnUnit.HeaderText = "Unidade";
            this.ColumnUnit.Name = "ColumnUnit";
            this.ColumnUnit.Width = 72;
            // 
            // ManufacturerList
            // 
            this.ManufacturerList.HeaderText = "Fabricante";
            this.ManufacturerList.Name = "ManufacturerList";
            // 
            // CodeList
            // 
            this.CodeList.HeaderText = "Código";
            this.CodeList.Name = "CodeList";
            // 
            // ObsList
            // 
            this.ObsList.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.ObsList.HeaderText = "Observações";
            this.ObsList.Name = "ObsList";
            this.ObsList.Width = 95;
            // 
            // btMatList
            // 
            this.btMatList.Location = new System.Drawing.Point(34, 409);
            this.btMatList.Name = "btMatList";
            this.btMatList.Size = new System.Drawing.Size(75, 23);
            this.btMatList.TabIndex = 2;
            this.btMatList.Text = "Gerar";
            this.btMatList.UseVisualStyleBackColor = true;
            this.btMatList.Click += new System.EventHandler(this.btMatList_Click);
            // 
            // labelArqPath
            // 
            this.labelArqPath.AutoSize = true;
            this.labelArqPath.Location = new System.Drawing.Point(146, 29);
            this.labelArqPath.Name = "labelArqPath";
            this.labelArqPath.Size = new System.Drawing.Size(16, 13);
            this.labelArqPath.TabIndex = 1;
            this.labelArqPath.Text = "...";
            // 
            // lbArquivoSelect
            // 
            this.lbArquivoSelect.AutoSize = true;
            this.lbArquivoSelect.Location = new System.Drawing.Point(31, 29);
            this.lbArquivoSelect.Name = "lbArquivoSelect";
            this.lbArquivoSelect.Size = new System.Drawing.Size(109, 13);
            this.lbArquivoSelect.TabIndex = 0;
            this.lbArquivoSelect.Text = "Arquivo selecionado: ";
            // 
            // tabPageOpMaterial
            // 
            this.tabPageOpMaterial.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageOpMaterial.Controls.Add(this.tbRestart);
            this.tabPageOpMaterial.Controls.Add(this.btSave);
            this.tabPageOpMaterial.Controls.Add(this.groupBox3);
            this.tabPageOpMaterial.Controls.Add(this.groupBox2);
            this.tabPageOpMaterial.Location = new System.Drawing.Point(4, 22);
            this.tabPageOpMaterial.Name = "tabPageOpMaterial";
            this.tabPageOpMaterial.Size = new System.Drawing.Size(697, 455);
            this.tabPageOpMaterial.TabIndex = 3;
            this.tabPageOpMaterial.Text = "Opções Lista de material";
            // 
            // tbRestart
            // 
            this.tbRestart.Location = new System.Drawing.Point(601, 417);
            this.tbRestart.Name = "tbRestart";
            this.tbRestart.Size = new System.Drawing.Size(75, 23);
            this.tbRestart.TabIndex = 8;
            this.tbRestart.Text = "Restaurar";
            this.tbRestart.UseVisualStyleBackColor = true;
            this.tbRestart.Click += new System.EventHandler(this.tbRestart_Click);
            // 
            // btSave
            // 
            this.btSave.Location = new System.Drawing.Point(17, 417);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(75, 23);
            this.btSave.TabIndex = 7;
            this.btSave.Text = "Salvar";
            this.btSave.UseVisualStyleBackColor = true;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.tbSPDCode);
            this.groupBox3.Controls.Add(this.label19);
            this.groupBox3.Controls.Add(this.tbTerminalCode);
            this.groupBox3.Controls.Add(this.tbRCDCode);
            this.groupBox3.Controls.Add(this.tbBreakerCode);
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Controls.Add(this.label14);
            this.groupBox3.Controls.Add(this.label15);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(17, 175);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(627, 147);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Código";
            // 
            // tbSPDCode
            // 
            this.tbSPDCode.Location = new System.Drawing.Point(120, 113);
            this.tbSPDCode.Name = "tbSPDCode";
            this.tbSPDCode.Size = new System.Drawing.Size(486, 22);
            this.tbSPDCode.TabIndex = 11;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label19.Location = new System.Drawing.Point(82, 118);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(32, 13);
            this.label19.TabIndex = 10;
            this.label19.Text = "DPS:";
            // 
            // tbTerminalCode
            // 
            this.tbTerminalCode.Location = new System.Drawing.Point(120, 84);
            this.tbTerminalCode.Name = "tbTerminalCode";
            this.tbTerminalCode.Size = new System.Drawing.Size(486, 22);
            this.tbTerminalCode.TabIndex = 5;
            // 
            // tbRCDCode
            // 
            this.tbRCDCode.Location = new System.Drawing.Point(120, 56);
            this.tbRCDCode.Name = "tbRCDCode";
            this.tbRCDCode.Size = new System.Drawing.Size(486, 22);
            this.tbRCDCode.TabIndex = 4;
            // 
            // tbBreakerCode
            // 
            this.tbBreakerCode.Location = new System.Drawing.Point(120, 28);
            this.tbBreakerCode.Name = "tbBreakerCode";
            this.tbBreakerCode.Size = new System.Drawing.Size(486, 22);
            this.tbBreakerCode.TabIndex = 3;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(64, 89);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(50, 13);
            this.label13.TabIndex = 2;
            this.label13.Text = "Terminal:";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(20, 61);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(95, 13);
            this.label14.TabIndex = 1;
            this.label14.Text = "Disjuntor Residual:";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(63, 33);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(51, 13);
            this.label15.TabIndex = 0;
            this.label15.Text = "Disjuntor:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.tbSPDManuf);
            this.groupBox2.Controls.Add(this.label16);
            this.groupBox2.Controls.Add(this.tbTerminalManuf);
            this.groupBox2.Controls.Add(this.tbRCDManuf);
            this.groupBox2.Controls.Add(this.tbBreakerManuf);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(17, 21);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(627, 148);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Fabricante";
            // 
            // tbSPDManuf
            // 
            this.tbSPDManuf.Location = new System.Drawing.Point(120, 113);
            this.tbSPDManuf.Name = "tbSPDManuf";
            this.tbSPDManuf.Size = new System.Drawing.Size(486, 22);
            this.tbSPDManuf.TabIndex = 7;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(82, 118);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(32, 13);
            this.label16.TabIndex = 6;
            this.label16.Text = "DPS:";
            // 
            // tbTerminalManuf
            // 
            this.tbTerminalManuf.Location = new System.Drawing.Point(120, 84);
            this.tbTerminalManuf.Name = "tbTerminalManuf";
            this.tbTerminalManuf.Size = new System.Drawing.Size(486, 22);
            this.tbTerminalManuf.TabIndex = 5;
            // 
            // tbRCDManuf
            // 
            this.tbRCDManuf.Location = new System.Drawing.Point(120, 56);
            this.tbRCDManuf.Name = "tbRCDManuf";
            this.tbRCDManuf.Size = new System.Drawing.Size(486, 22);
            this.tbRCDManuf.TabIndex = 4;
            // 
            // tbBreakerManuf
            // 
            this.tbBreakerManuf.Location = new System.Drawing.Point(120, 28);
            this.tbBreakerManuf.Name = "tbBreakerManuf";
            this.tbBreakerManuf.Size = new System.Drawing.Size(486, 22);
            this.tbBreakerManuf.TabIndex = 3;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(64, 89);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(50, 13);
            this.label12.TabIndex = 2;
            this.label12.Text = "Terminal:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(20, 61);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(95, 13);
            this.label11.TabIndex = 1;
            this.label11.Text = "Disjuntor Residual:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(63, 33);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(51, 13);
            this.label10.TabIndex = 0;
            this.label10.Text = "Disjuntor:";
            // 
            // tabPageLoads
            // 
            this.tabPageLoads.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageLoads.Controls.Add(this.btLoadFileFromExcel);
            this.tabPageLoads.Controls.Add(this.btLoadToCad);
            this.tabPageLoads.Controls.Add(this.dgvLoads);
            this.tabPageLoads.Location = new System.Drawing.Point(4, 22);
            this.tabPageLoads.Name = "tabPageLoads";
            this.tabPageLoads.Size = new System.Drawing.Size(697, 455);
            this.tabPageLoads.TabIndex = 7;
            this.tabPageLoads.Text = "Cargas";
            // 
            // btLoadFileFromExcel
            // 
            this.btLoadFileFromExcel.Location = new System.Drawing.Point(14, 416);
            this.btLoadFileFromExcel.Name = "btLoadFileFromExcel";
            this.btLoadFileFromExcel.Size = new System.Drawing.Size(75, 23);
            this.btLoadFileFromExcel.TabIndex = 2;
            this.btLoadFileFromExcel.Text = "Carregar";
            this.btLoadFileFromExcel.UseVisualStyleBackColor = true;
            this.btLoadFileFromExcel.Click += new System.EventHandler(this.btLoadFileFromExcel_Click);
            // 
            // btLoadToCad
            // 
            this.btLoadToCad.Location = new System.Drawing.Point(95, 416);
            this.btLoadToCad.Name = "btLoadToCad";
            this.btLoadToCad.Size = new System.Drawing.Size(75, 23);
            this.btLoadToCad.TabIndex = 1;
            this.btLoadToCad.Text = "Inserir";
            this.btLoadToCad.UseVisualStyleBackColor = true;
            this.btLoadToCad.Click += new System.EventHandler(this.btLoadToCad_Click);
            // 
            // dgvLoads
            // 
            this.dgvLoads.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvLoads.Location = new System.Drawing.Point(14, 33);
            this.dgvLoads.Name = "dgvLoads";
            this.dgvLoads.Size = new System.Drawing.Size(666, 360);
            this.dgvLoads.TabIndex = 0;
            // 
            // tabPagePanelMount
            // 
            this.tabPagePanelMount.BackColor = System.Drawing.SystemColors.Control;
            this.tabPagePanelMount.Controls.Add(this.groupBox12);
            this.tabPagePanelMount.Controls.Add(this.btMountPanel);
            this.tabPagePanelMount.Location = new System.Drawing.Point(4, 22);
            this.tabPagePanelMount.Name = "tabPagePanelMount";
            this.tabPagePanelMount.Size = new System.Drawing.Size(697, 455);
            this.tabPagePanelMount.TabIndex = 6;
            this.tabPagePanelMount.Text = "Montagem";
            // 
            // groupBox12
            // 
            this.groupBox12.Controls.Add(this.groupBox13);
            this.groupBox12.Controls.Add(this.label37);
            this.groupBox12.Controls.Add(this.comboBox1);
            this.groupBox12.Location = new System.Drawing.Point(23, 16);
            this.groupBox12.Name = "groupBox12";
            this.groupBox12.Size = new System.Drawing.Size(656, 366);
            this.groupBox12.TabIndex = 3;
            this.groupBox12.TabStop = false;
            // 
            // groupBox13
            // 
            this.groupBox13.Location = new System.Drawing.Point(9, 51);
            this.groupBox13.Name = "groupBox13";
            this.groupBox13.Size = new System.Drawing.Size(377, 100);
            this.groupBox13.TabIndex = 4;
            this.groupBox13.TabStop = false;
            this.groupBox13.Text = "Barramento de neutro";
            // 
            // label37
            // 
            this.label37.AutoSize = true;
            this.label37.Location = new System.Drawing.Point(6, 16);
            this.label37.Name = "label37";
            this.label37.Size = new System.Drawing.Size(108, 13);
            this.label37.TabIndex = 1;
            this.label37.Text = "Bloco de distribuição:";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "7 LIGAÇÕES",
            "11 LIGAÇÕES",
            "15 LIGAÇÕES"});
            this.comboBox1.Location = new System.Drawing.Point(120, 13);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 2;
            // 
            // btMountPanel
            // 
            this.btMountPanel.Location = new System.Drawing.Point(23, 415);
            this.btMountPanel.Name = "btMountPanel";
            this.btMountPanel.Size = new System.Drawing.Size(100, 23);
            this.btMountPanel.TabIndex = 0;
            this.btMountPanel.Text = "Gerar Montagem";
            this.btMountPanel.UseVisualStyleBackColor = true;
            this.btMountPanel.Click += new System.EventHandler(this.btMountPanel_Click);
            // 
            // tbAbout
            // 
            this.tbAbout.BackColor = System.Drawing.SystemColors.Control;
            this.tbAbout.Controls.Add(this.btManual);
            this.tbAbout.Controls.Add(this.groupBox10);
            this.tbAbout.Controls.Add(this.lbVersion);
            this.tbAbout.Controls.Add(this.label32);
            this.tbAbout.Controls.Add(this.label31);
            this.tbAbout.Controls.Add(this.groupBox9);
            this.tbAbout.Controls.Add(this.groupBox8);
            this.tbAbout.Controls.Add(this.label28);
            this.tbAbout.Location = new System.Drawing.Point(4, 22);
            this.tbAbout.Name = "tbAbout";
            this.tbAbout.Size = new System.Drawing.Size(697, 455);
            this.tbAbout.TabIndex = 4;
            this.tbAbout.Text = "Sobre";
            // 
            // btManual
            // 
            this.btManual.Location = new System.Drawing.Point(21, 410);
            this.btManual.Name = "btManual";
            this.btManual.Size = new System.Drawing.Size(91, 23);
            this.btManual.TabIndex = 31;
            this.btManual.Text = "Abrir Manual";
            this.btManual.UseVisualStyleBackColor = true;
            this.btManual.Click += new System.EventHandler(this.btManual_Click);
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.linkLabel3);
            this.groupBox10.Controls.Add(this.linkLabel4);
            this.groupBox10.Location = new System.Drawing.Point(21, 192);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(241, 67);
            this.groupBox10.TabIndex = 30;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "GitHub";
            // 
            // linkLabel3
            // 
            this.linkLabel3.AutoSize = true;
            this.linkLabel3.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel3.Location = new System.Drawing.Point(6, 16);
            this.linkLabel3.Name = "linkLabel3";
            this.linkLabel3.Size = new System.Drawing.Size(154, 13);
            this.linkLabel3.TabIndex = 29;
            this.linkLabel3.TabStop = true;
            this.linkLabel3.Text = "https://github.com/gustavohsz";
            // 
            // linkLabel4
            // 
            this.linkLabel4.AutoSize = true;
            this.linkLabel4.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel4.Location = new System.Drawing.Point(6, 42);
            this.linkLabel4.Name = "linkLabel4";
            this.linkLabel4.Size = new System.Drawing.Size(172, 13);
            this.linkLabel4.TabIndex = 28;
            this.linkLabel4.TabStop = true;
            this.linkLabel4.Text = "https://github.com/sergiopetrovcic";
            // 
            // lbVersion
            // 
            this.lbVersion.AutoSize = true;
            this.lbVersion.Location = new System.Drawing.Point(93, 297);
            this.lbVersion.Name = "lbVersion";
            this.lbVersion.Size = new System.Drawing.Size(41, 13);
            this.lbVersion.TabIndex = 30;
            this.lbVersion.Text = "version";
            // 
            // label32
            // 
            this.label32.AutoSize = true;
            this.label32.Location = new System.Drawing.Point(18, 297);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(69, 13);
            this.label32.TabIndex = 29;
            this.label32.Text = "Versão atual:";
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(18, 275);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(244, 13);
            this.label31.TabIndex = 28;
            this.label31.Text = "Instituto Federal de Santa Catarina - Campus Itajaí";
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.linkLabel2);
            this.groupBox9.Controls.Add(this.linkLabel1);
            this.groupBox9.Location = new System.Drawing.Point(21, 119);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(241, 67);
            this.groupBox9.TabIndex = 27;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Email";
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel2.Location = new System.Drawing.Point(6, 16);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(153, 13);
            this.linkLabel2.TabIndex = 29;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "gustavo.hsz@aluno.ifsc.edu.br";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel1.Location = new System.Drawing.Point(6, 42);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(142, 13);
            this.linkLabel1.TabIndex = 28;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "sergio.petrovcic@ifsc.edu.br";
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.label30);
            this.groupBox8.Controls.Add(this.label29);
            this.groupBox8.Location = new System.Drawing.Point(21, 50);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(241, 63);
            this.groupBox8.TabIndex = 26;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Desenvolvedores:";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Location = new System.Drawing.Point(6, 38);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(220, 13);
            this.label30.TabIndex = 1;
            this.label30.Text = "Prof. Dr. Sergio Augusto Bitencourt Petrovcic";
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(6, 16);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(155, 13);
            this.label29.TabIndex = 0;
            this.label29.Text = "Gustavo Henrique Silva Zanetti";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label28.Location = new System.Drawing.Point(17, 13);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(164, 24);
            this.label28.TabIndex = 25;
            this.label28.Text = "Sheet To CAD ®️";
            // 
            // InsertForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(720, 497);
            this.Controls.Add(this.tabControl1);
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InsertForm";
            this.ShowIcon = false;
            this.Text = "STC - Sheet to CAD";
            this.tabControl1.ResumeLayout(false);
            this.tabPageMain.ResumeLayout(false);
            this.tabPageMain.PerformLayout();
            this.gbCoordinates.ResumeLayout(false);
            this.gbCoordinates.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPageConfig.ResumeLayout(false);
            this.tabPageConfig.PerformLayout();
            this.groupBox11.ResumeLayout(false);
            this.groupBox11.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.tabPageMaterialList.ResumeLayout(false);
            this.tabPageMaterialList.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMaterialList)).EndInit();
            this.tabPageOpMaterial.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabPageLoads.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvLoads)).EndInit();
            this.tabPagePanelMount.ResumeLayout(false);
            this.groupBox12.ResumeLayout(false);
            this.groupBox12.PerformLayout();
            this.tbAbout.ResumeLayout(false);
            this.tbAbout.PerformLayout();
            this.groupBox10.ResumeLayout(false);
            this.groupBox10.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.HelpProvider helpProvider1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageMain;
        private System.Windows.Forms.TabPage tabPageMaterialList;
        private System.Windows.Forms.GroupBox gbCoordinates;
        private System.Windows.Forms.TextBox txtCoordY;
        private System.Windows.Forms.TextBox txtCoordX;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox TBBus;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox CBGrounding;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox CBTube;
        private System.Windows.Forms.TextBox TBSource;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox CBVoltage;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.Label lbInfo;
        private System.Windows.Forms.RadioButton rbScreen;
        private System.Windows.Forms.RadioButton rbCoordinates;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btFileSearch;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btOk;
        private System.Windows.Forms.TabPage tabPageConfig;
        private System.Windows.Forms.Label labelArqPath;
        private System.Windows.Forms.Label lbArquivoSelect;
        private System.Windows.Forms.Button btMatList;
        private System.Windows.Forms.DataGridView dgvMaterialList;
        private System.Windows.Forms.TabPage tabPageOpMaterial;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox tbTerminalCode;
        private System.Windows.Forms.TextBox tbRCDCode;
        private System.Windows.Forms.TextBox tbBreakerCode;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox tbTerminalManuf;
        private System.Windows.Forms.TextBox tbRCDManuf;
        private System.Windows.Forms.TextBox tbBreakerManuf;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label lbNotes;
        private System.Windows.Forms.DataGridViewTextBoxColumn DescriptionList;
        private System.Windows.Forms.DataGridViewTextBoxColumn QuantityList;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnUnit;
        private System.Windows.Forms.DataGridViewTextBoxColumn ManufacturerList;
        private System.Windows.Forms.DataGridViewTextBoxColumn CodeList;
        private System.Windows.Forms.DataGridViewTextBoxColumn ObsList;
        private System.Windows.Forms.TextBox tbSPDManuf;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox tbSPDCode;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Button btSave;
        private System.Windows.Forms.Button tbRestart;
        private System.Windows.Forms.Button btCancelConfig;
        private System.Windows.Forms.Button btApplyConfig;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.ComboBox CBFontName;
        private System.Windows.Forms.TextBox TBFontSize;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox TBScale;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.TextBox TBGroundLayer;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.TextBox TBWiringLayer;
        private System.Windows.Forms.TextBox TBNeutralLayer;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.TextBox TBCircuitLayer;
        private System.Windows.Forms.TextBox TBBusLayer;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.ComboBox cbSheets;
        private System.Windows.Forms.Button btExcelExport;
        private System.Windows.Forms.TabPage tbAbout;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.Label label32;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Label label30;
        private System.Windows.Forms.GroupBox groupBox10;
        private System.Windows.Forms.LinkLabel linkLabel3;
        private System.Windows.Forms.LinkLabel linkLabel4;
        private System.Windows.Forms.Label lbVersion;
        private System.Windows.Forms.Button btCADExport;
        private System.Windows.Forms.GroupBox groupBox11;
        private System.Windows.Forms.TextBox tbPhaseC;
        private System.Windows.Forms.TextBox tbPhaseB;
        private System.Windows.Forms.TextBox tbPhaseA;
        private System.Windows.Forms.Label label36;
        private System.Windows.Forms.Label label35;
        private System.Windows.Forms.Label label34;
        private System.Windows.Forms.TabPage tabPagePanelMount;
        private System.Windows.Forms.Button btMountPanel;
        private System.Windows.Forms.TabPage tabPageLoads;
        private System.Windows.Forms.DataGridView dgvLoads;
        private System.Windows.Forms.Button btLoadToCad;
        private System.Windows.Forms.GroupBox groupBox12;
        private System.Windows.Forms.Label label37;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.GroupBox groupBox13;
        private System.Windows.Forms.Button btLoadFileFromExcel;
        private System.Windows.Forms.Button btManual;
    }
}