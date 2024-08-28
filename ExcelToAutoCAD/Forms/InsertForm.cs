using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.IO;
using Newtonsoft.Json;
using ExcelToAutoCAD.Save;
using ExcelToAutoCAD.DataSaves;
using Microsoft.Office.Interop.Excel;
using ExcelToAutoCAD.Entities;
using ExcelToAutoCAD.PrintMgr;
using ExcelToAutoCAD.Menu;
using ExcelToAutoCAD.Panel;
using Newtonsoft.Json.Linq;

namespace ExcelToAutoCAD
{

    public partial class InsertForm : Form
    {
        Point3d insPt;
        double _posX, _posY;
        string pathFile;
        private ExcelToAutoCAD _exTADCAD;
        private string sheetName = "";

        private const string saveFile = @"saveFile.json";
        private PrintDraw printDraw = new PrintDraw();
        private LoadDrawingPaper loadDrawingPaper = new LoadDrawingPaper();
        private ExcelAccess ReadExcel = new ExcelAccess();

        public string boardName = "QD";

        public InsertForm(ExcelToAutoCAD exTACAD)
        {

            InitializeComponent();
            _exTADCAD = exTACAD;

            tabControl1 = new TabControl();
                        

            LoadTexts();

            LoadSave();

            string version = "V.1.115";
            lbVersion.Text = version;
            this.Text = "STC - Sheet to CAD " + version;

            gbCoordinates.Enabled = false;
            rbScreen.Checked = true;

        }

        private string CopyFile(string fileName)
        {
            string tempFile = Path.GetTempFileName();
            File.Copy(fileName, tempFile, true);
            return tempFile;
        }
                

        public void LoadSave()
        {
            if (File.Exists(saveFile))
            {
                string json = File.ReadAllText(saveFile);
                DataMain data = JsonConvert.DeserializeObject<DataMain>(json);

                if (data.dataMaterialList != null)
                {
                    tbBreakerManuf.Text = data.dataMaterialList.dataBreakerManuf ?? "";
                    tbRCDManuf.Text = data.dataMaterialList.dataRCDManuf ?? "";
                    tbTerminalManuf.Text = data.dataMaterialList.dataTerminalManuf ?? "";
                    tbSPDManuf.Text = data.dataMaterialList.dataSPDManuf ?? "";
                    tbBreakerCode.Text = data.dataMaterialList.dataBreakerCode ?? "";
                    tbRCDCode.Text = data.dataMaterialList.dataRCDCode ?? "";
                    tbTerminalCode.Text = data.dataMaterialList.dataTerminalCode ?? "";
                    tbSPDCode.Text = data.dataMaterialList.dataSPDCode ?? "";
                }
                if (data.dataInsertFile != null)
                {
                    CBVoltage.Text = data.dataInsertFile.dataVoltage ?? "";
                    CBTube.Text = data.dataInsertFile.dataTube ?? "";
                    TBSource.Text = data.dataInsertFile.dataSource ?? "";
                    CBGrounding.Text = data.dataInsertFile.dataGrounding ?? "";
                    TBBus.Text = data.dataInsertFile.dataBus ?? "";
                }
                if (data.dataConfig != null)
                {
                    TBScale.Text = data.dataConfig.dataScale ?? "";
                    TBCircuitLayer.Text = data.dataConfig.dataLCircuits ?? "";
                    TBBusLayer.Text = data.dataConfig.dataLBus ?? "";
                    TBNeutralLayer.Text = data.dataConfig.dataLNeutral ?? "";
                    TBGroundLayer.Text = data.dataConfig.dataLGround ?? "";
                    TBWiringLayer.Text = data.dataConfig.dataLWiring ?? "";
                    CBFontName.Text = data.dataConfig.dataTType ?? "";
                    TBFontSize.Text = data.dataConfig.dataTSize ?? "";
                    tbPhaseA.Text = data.dataConfig.dataPhaseA ?? _exTADCAD.PhaseLaterA;
                    tbPhaseB.Text = data.dataConfig.dataPhaseB ?? _exTADCAD.PhaseLaterB;
                    tbPhaseC.Text = data.dataConfig.dataPhaseC ?? _exTADCAD.PhaseLaterC;
                }
            }
        }

        public void LoadTexts()
        {
            //geral
            TBSource.Text = _exTADCAD.Source;
            CBVoltage.Text = _exTADCAD.Voltage;
            CBTube.Text = _exTADCAD.Tube;
            CBGrounding.Text = _exTADCAD.GroundingSystem;
            TBBus.Text = _exTADCAD.BusSystem;

            //lista de material
            tbBreakerManuf.Text = "Siemens";
            tbRCDManuf.Text = "Siemens";
            tbTerminalManuf.Text = "Phoenix Contact";
            tbSPDManuf.Text = "Clamper";

            tbBreakerCode.Text = "5SL3";
            tbRCDCode.Text = "5SV";
            tbTerminalCode.Text = "-";
            tbSPDCode.Text = "014295";

            //configurações
            TBCircuitLayer.Text = "STC - Circuitos";
            TBNeutralLayer.Text = "STC - Neutro";
            TBGroundLayer.Text = "STC - Terra";
            TBBusLayer.Text = "STC - Barramento";
            TBWiringLayer.Text = "STC - Fiação";
            TBScale.Text = _exTADCAD.scaleFactor.ToString();
            tbPhaseA.Text = "A";
            tbPhaseB.Text = "B";
            tbPhaseC.Text = "C";

            foreach (FontFamily font in System.Drawing.FontFamily.Families) //Tipo de fontes
            {
                CBFontName.Items.Add(font.Name);
            }

            if (CBFontName.Items.Contains(_exTADCAD.FontName.ToString()))
            {
                CBFontName.SelectedItem = _exTADCAD.FontName.ToString();
            }
            else
            {
                CBFontName.Items.Add(_exTADCAD.FontName.ToString());
                CBFontName.SelectedItem = _exTADCAD.FontName.ToString();
            }

            if (TBFontSize.Text.Contains(_exTADCAD.FontSize.ToString())) //tamanho da fonte
            {
                TBFontSize.Text = _exTADCAD.FontSize.ToString();
            }
            else
            {
                TBFontSize.Text = _exTADCAD.FontSize.ToString();
            }


            List<string> list = new List<string>()
                {
                    "Circuito",
                    "Descrição",
                    "Potência",
                    "Fase A",
                    "Fase B",
                    "Fase C",
                    "Condutor",
                    "Disjuntor",
                    "IDR"
                };

            foreach (string column in list)
            {
                dgvLoads.Columns.Add(column, column);
            }

        }

        public void LoadSheet()
        {
            ExcelAccess ReadExcel = new ExcelAccess();

            ReadExcel.ReadColumn(pathFile, sheetName);
            ExcelAccess.Row row;

            row = ReadExcel.rows[1];                        
            boardName = row.Title;

            for (int i = 0; i <ReadExcel.rows.Count; i++)
            {
                row = ReadExcel.rows[i];
                dgvLoads.Refresh();

                if (row.Circuit != null && row.Circuit != "")
                {

                    dgvLoads.Rows.Add();

                    dgvLoads.Rows[i].Cells["Circuito"].Value = row.Circuit;
                    dgvLoads.Rows[i].Cells["Descrição"].Value = row.Description;
                    dgvLoads.Rows[i].Cells["Potência"].Value = row.Power;

                    if (decimal.TryParse(row.PhaseA, out decimal faseA))
                        dgvLoads.Rows[i].Cells["Fase A"].Value = Math.Round(faseA);
                    if (decimal.TryParse(row.PhaseB, out decimal faseB))
                        dgvLoads.Rows[i].Cells["Fase B"].Value = Math.Round(faseB);
                    if (decimal.TryParse(row.PhaseC, out decimal faseC))
                        dgvLoads.Rows[i].Cells["Fase C"].Value = Math.Round(faseC);

                    dgvLoads.Rows[i].Cells["Condutor"].Value = row.ConductorSection;                                      
                    
                    
                    if(row.CBPolesQuantity == "UNIPOLAR")
                        dgvLoads.Rows[i].Cells["Disjuntor"].Value = "Curva " + row.CBCurve + " - " + "1x" +  row.CBCurrent;
                    if (row.CBPolesQuantity == "BIPOLAR")
                        dgvLoads.Rows[i].Cells["Disjuntor"].Value = "Curva " + row.CBCurve + " - " + "2x" + row.CBCurrent;
                    if (row.CBPolesQuantity == "TRIPOLAR")
                        dgvLoads.Rows[i].Cells["Disjuntor"].Value = "Curva " + row.CBCurve + " - " + "3x" + row.CBCurrent;
                                        

                    if (row.RCDNumbering != "0" && row.RCDCurrent != "")
                    {
                        
                        if (row.RCDPolesQuant == "BIPOLAR")
                        {
                            dgvLoads.Rows[i].Cells["IDR"].Value = "IDR" + row.RCDNumbering + "-2P-" + row.RCDCurrent;
                        }
                        else
                        {
                            dgvLoads.Rows[i].Cells["IDR"].Value = "IDR" + row.RCDNumbering + "-4P-" + row.RCDCurrent;
                        }
                    }
                    dgvLoads.Refresh();
                }
                
            }
            for(int j =0; j< ReadExcel.others.Count; j++)
            {
                row = ReadExcel.others[j];
                if (row.Circuit == "ENTRADA DE ENERGIA")
                {
                   int idx = dgvLoads.Rows.Add();

                    dgvLoads.Rows[idx].Cells["Circuito"].Value = row.Circuit;
                    dgvLoads.Rows[idx].Cells["Descrição"].Value = row.Description;
                    dgvLoads.Rows[idx].Cells["Potência"].Value = row.Power;

                    if (decimal.TryParse(row.PhaseA, out decimal faseA))
                        dgvLoads.Rows[idx].Cells["Fase A"].Value = Math.Round(faseA);
                    if (decimal.TryParse(row.PhaseB, out decimal faseB))
                        dgvLoads.Rows[idx].Cells["Fase B"].Value = Math.Round(faseB);
                    if (decimal.TryParse(row.PhaseC, out decimal faseC))
                        dgvLoads.Rows[idx].Cells["Fase C"].Value = Math.Round(faseC);

                    dgvLoads.Rows[idx].Cells["Condutor"].Value = row.ConductorSection;

                    if (row.CBPolesQuantity == "UNIPOLAR")
                        dgvLoads.Rows[idx].Cells["Disjuntor"].Value = "Curva " + row.CBCurve + " - " + "1x" + row.CBCurrent;
                    if (row.CBPolesQuantity == "BIPOLAR")
                        dgvLoads.Rows[idx].Cells["Disjuntor"].Value = "Curva " + row.CBCurve + " - " + "2x" + row.CBCurrent;
                    if (row.CBPolesQuantity == "TRIPOLAR")
                        dgvLoads.Rows[idx].Cells["Disjuntor"].Value = "Curva " + row.CBCurve + " - " + "3x" + row.CBCurrent;

                    dgvLoads.Refresh();
                }

            }
            

        }

        public void SaveConfig()
        {
            List<DataMain> dataList = new List<DataMain>();

            DataMain dataMain = new DataMain
            {
                dataMaterialList = new DataMaterialList
                {
                    dataBreakerManuf = tbBreakerManuf.Text,
                    dataRCDManuf = tbRCDManuf.Text,
                    dataTerminalManuf = tbTerminalManuf.Text,
                    dataSPDManuf = tbSPDManuf.Text,
                    dataBreakerCode = tbBreakerCode.Text,
                    dataRCDCode = tbRCDCode.Text,
                    dataTerminalCode = tbTerminalCode.Text,
                    dataSPDCode = tbSPDCode.Text
                },

                dataInsertFile = new DataInsertFile
                {
                    dataVoltage = CBVoltage.Text,
                    dataTube = CBTube.Text,
                    dataSource = TBSource.Text,
                    dataGrounding = CBGrounding.Text,
                    dataBus = TBBus.Text

                },
                dataConfig = new DataConfig
                {
                    dataScale = TBScale.Text,
                    dataLCircuits = TBCircuitLayer.Text,
                    dataLBus = TBBusLayer.Text,
                    dataLNeutral = TBNeutralLayer.Text,
                    dataLGround = TBGroundLayer.Text,
                    dataLWiring = TBWiringLayer.Text,
                    dataTType = CBFontName.Text,
                    dataTSize = TBFontSize.Text,
                    dataPhaseA = tbPhaseA.Text,
                    dataPhaseB = tbPhaseB.Text,
                    dataPhaseC = tbPhaseC.Text
                }
            };
            string json = JsonConvert.SerializeObject(dataMain, Formatting.None);
            File.WriteAllText(saveFile, json);

            MessageBox.Show("Informações salvas!", "Salvamento", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void ExportToExcel(DataGridView dgv)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Arquivos do Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*";
                saveFileDialog.Title = "Salvar arquivo do excel";
                saveFileDialog.FileName = "Lista de material.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = saveFileDialog.FileName;
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = true;

                    Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                    for (int i = 1; i < dgv.Columns.Count + 1; i++) //copia o cabeçalho
                    {
                        worksheet.Cells[1, i] = dgv.Columns[i - 1].HeaderText;
                    }
                    for (int i = 0; i < dgv.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            if (dgv.Rows[i].Cells[j].Value != null)
                                worksheet.Cells[i + 2, j + 1] = dgv.Rows[i].Cells[j].Value.ToString();
                            else
                                worksheet.Cells[i + 2, j + 1] = "";
                        }
                    }
                    workbook.SaveAs(folderPath);

                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                    workbook = null;
                    worksheet = null;
                    excelApp = null;
                    GC.Collect();

                    MessageBox.Show("Exportação concluída", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Erro ao exportar para o excel: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btFileSearch_Click_1(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.Filter = "Arquivos do Excel (*.xlsx, *.xlsm)|*.xlsx;*.xlsm|Todos os arquivos (*.*)|*.*";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = dlg.FileName;
                    pathFile = dlg.FileName;
                    labelArqPath.Text = dlg.FileName;
                }
                if (!string.IsNullOrEmpty(pathFile))
                {
                    string fileName = pathFile;
                    string tempFile = CopyFile(fileName);

                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Workbook workbook = excelApp.Workbooks.Open(tempFile);

                    List<string> worksheetsNames = new List<string>();

                    foreach (Worksheet sheet in workbook.Sheets)
                    {
                        if (sheet.Name != "DEFAULT")
                            worksheetsNames.Add(sheet.Name);
                    }
                    workbook.Close();
                    excelApp.Quit();

                    cbSheets.DataSource = worksheetsNames;

                    

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro: " + ex, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void btCancel_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void rbCoordinates_CheckedChanged_1(object sender, EventArgs e)
        {
            gbCoordinates.Enabled = true;
        }

        private void rbScreen_CheckedChanged_1(object sender, EventArgs e)
        {
            gbCoordinates.Enabled = false;
        }

        private void btOk_Click_1(object sender, EventArgs e)
        {
            _exTADCAD.GroundingSystem = CBGrounding.Text;
            _exTADCAD.BusSystem = TBBus.Text;
            _exTADCAD.Voltage = CBVoltage.Text;
            _exTADCAD.Source = TBSource.Text;
            _exTADCAD.Tube = CBTube.Text;

            if (gbCoordinates.Enabled == true)
            {
                bool isValid = ValidateEntry(txtCoordX);
                if (isValid == false)
                {
                    MessageBox.Show("Coordenada X com entrada inválida. Por favor, entre com um valor válido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtCoordX.Focus();
                    return;
                }
                isValid = ValidateEntry(txtCoordY);
                if (isValid == false)
                {
                    MessageBox.Show("Coordenada Y com entrada inválida. Por favor, entre com um valor válido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtCoordY.Focus();
                    return;
                }
                _posX = double.Parse(txtCoordX.Text.Trim());
                _posY = double.Parse(txtCoordY.Text.Trim());
                insPt = new Point3d(_posX, _posY, 0);
            }

            if (rbScreen.Checked)
            {
                this.Hide();
                Editor edt = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                PromptPointOptions ppo = new PromptPointOptions("Escolha o ponto inicial: ");
                PromptPointResult ppr = edt.GetPoint(ppo);
                insPt = ppr.Value;
                this.Show();
            }


            if (!string.IsNullOrWhiteSpace(pathFile))
            {
                _exTADCAD.ExcelToCAD(insPt, pathFile, sheetName);
            }
            else
            {
                MessageBox.Show("Selecione um arquivo!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtFilePath.Focus();
                return;
            }

            
            lbInfo.Text = $"Unifilar criado com sucesso! Tempo de execução: {_exTADCAD.totalTime} ms";
            lbInfo.ForeColor = Color.Green;

            this.DialogResult = DialogResult.OK;
        }

        private void btMatList_Click(object sender, EventArgs e) //aba de lista de material
        {
            dgvMaterialList.Rows.Clear();

            if (!string.IsNullOrEmpty(pathFile))
            {
                ListaMaterial listaMaterial = new ListaMaterial(pathFile, sheetName);

                Dictionary<string, int> breakerCounts = listaMaterial.GetBreakersCount();
                Dictionary<string, int> terminalCounts = listaMaterial.GetTerminalsCount();
                Dictionary<string, int> rcdCounts = listaMaterial.GetRCDCount();
                Dictionary<string, int> spdCounts = listaMaterial.GetSPD();

                int totalBreakers = listaMaterial.GetPanelSize();
                int reserveBreaker = 0;
                string unitList = "Un.";


                foreach (var item in breakerCounts)
                {
                    dgvMaterialList.Rows.Add(item.Key, item.Value, unitList, tbBreakerManuf.Text, tbBreakerCode.Text);

                }
                foreach (var item in terminalCounts)
                {
                    dgvMaterialList.Rows.Add(item.Key, 4 * item.Value, unitList, tbTerminalManuf.Text, tbTerminalCode.Text); //2 terminais disjuntor + 1 neutro + 1 terra
                }
                foreach (var item in rcdCounts)
                {
                    dgvMaterialList.Rows.Add(item.Key, item.Value, unitList, tbRCDManuf.Text, tbRCDCode.Text);
                }
                foreach (var item in spdCounts)
                {
                    dgvMaterialList.Rows.Add(item.Key, item.Value, unitList, tbSPDManuf.Text, tbSPDCode.Text);
                }
                //tabela 59 NBR 5410 (espaço reserva din)
                if (totalBreakers <= 6)
                    reserveBreaker += 2;

                else if (totalBreakers >= 7 && totalBreakers <= 12)
                    reserveBreaker += 3;

                else if (totalBreakers >= 13 && totalBreakers <= 30)
                    reserveBreaker += 4;

                else if (totalBreakers > 30)
                    reserveBreaker = (int)Math.Ceiling((0.15 * totalBreakers));

                totalBreakers += reserveBreaker;

                dgvMaterialList.Rows.Add("QUADRO MÍNIMO [DIN]", totalBreakers, "Peça", null, null, "Espaço reserva calculado: " + reserveBreaker + " [DIN]");


                lbNotes.Text = "Tamanho de quadro mínimo calculado\nconforme tabela 59 NBR 5410.";


            }
            else
            {
                MessageBox.Show("Selecione um arquivo!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtFilePath.Focus();
                return;
            }
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            SaveConfig();
        }

        private void tbRestart_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Isso irá restaurar todos os dados do formulário, deseja continuar?", "Restauração! ", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                LoadTexts();
            }
            else
            {
                return;
            }

        }

        private void btCancelConfig_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btApplyConfig_Click(object sender, EventArgs e)
        {
            _exTADCAD.CircuitsLayer = TBCircuitLayer.Text;
            _exTADCAD.NeutralLayer = TBNeutralLayer.Text;
            _exTADCAD.GroundLayer = TBGroundLayer.Text;
            _exTADCAD.BusLayer = TBBusLayer.Text;
            _exTADCAD.WiringLayer = TBWiringLayer.Text;

            _exTADCAD.scaleFactor = Convert.ToDouble(TBScale.Text);

            _exTADCAD.FontSize = Convert.ToDouble(TBFontSize.Text); //define o tamanho da fonte

            _exTADCAD.FontName = CBFontName.SelectedItem.ToString(); //define o tipo da fonte

            _exTADCAD.PhaseLaterA = tbPhaseA.Text;
            _exTADCAD.PhaseLaterB = tbPhaseB.Text;
            _exTADCAD.PhaseLaterC = tbPhaseC.Text;

            SaveConfig();
        }

        private bool ValidateEntry(System.Windows.Forms.TextBox tb) //método para validar entrada de dados nas coordenadas
        {
            bool isValid = false;
            double value;

            try
            {
                if (tb.Text.Trim() == "")
                {
                    lbInfo.Text = "Por favor, entre com um valor.";
                    lbInfo.ForeColor = Color.Red;
                }
                else
                {
                    value = double.Parse(tb.Text.Trim());
                    isValid = true;
                }
            }
            catch (Exception ex)
            {
                lbInfo.Text = "Entrada inválida: " + ex.Message;
                lbInfo.ForeColor = Color.Red;

            }

            return isValid;
        }

        private void btExcelExport_Click(object sender, EventArgs e)
        {
            ExportToExcel(dgvMaterialList);
        }

        private void btCADExport_Click(object sender, EventArgs e)
        {
            this.Hide();
            TableMaterialList tableMaterialList = new TableMaterialList();
            tableMaterialList.CreateTable(dgvMaterialList);
            this.Show();
        }
                

        private void btMountPanel_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(pathFile))
            {
                this.Hide();
                PanelMount panelMount = new PanelMount(pathFile, sheetName);
                panelMount.GeneratePanel();
                this.Show();
            }

        }

        private void btLoadToCad_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(pathFile))
            {
                this.Hide();
                TableLoads tableLoads = new TableLoads();
                tableLoads.CreateTable(dgvLoads, boardName);
                this.Show();
            }

            
        }

        private void btLoadFileFromExcel_Click(object sender, EventArgs e)
        {                      
            dgvLoads.Rows.Clear();
            dgvLoads.DataSource = null;
            dgvLoads.DataMember = null;
            
            dgvLoads.Refresh();

            if (!string.IsNullOrEmpty(sheetName))
                LoadSheet();
        }

        private void btManual_Click(object sender, EventArgs e)
        {
            OpenManual openManuak = new OpenManual();
            openManuak.Open();
        }

        private void cbSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            sheetName = cbSheets.SelectedItem.ToString();           

        }

    }
}
