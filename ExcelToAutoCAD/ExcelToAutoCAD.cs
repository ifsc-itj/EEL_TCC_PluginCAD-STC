using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;
using ExcelToAutoCAD.Entities;
using ExcelToAutoCAD.Menu;
using ExcelToAutoCAD.Panel;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using static ExcelToAutoCAD.ExcelAccess;
using System.Diagnostics;
using System.Threading;
using ExcelToAutoCAD.Forms;
using System.Threading.Tasks;


namespace ExcelToAutoCAD
{
    public class ExcelToAutoCAD : IExtensionApplication
    {

        public void Initialize()
        {
            /*bool var = true;

            while (var)
            {

                string relativeFolderPath = "config/install.txt";

                string currentDirectory = Environment.CurrentDirectory;

                string folderPath = Path.Combine(currentDirectory, relativeFolderPath);

                if (!File.Exists(folderPath))
                {
                    DateTime installDate = DateTime.Now;
                    File.WriteAllText(folderPath, installDate.ToString());
                }
                else
                {
                    string installDateStr = File.ReadAllText(folderPath);
                    DateTime installDate = DateTime.Parse(installDateStr);

                    TimeSpan diff = DateTime.Now - installDate;
                    var = false;
                    if (diff.TotalMinutes > 2)
                    {
                        System.Windows.Forms.MessageBox.Show("Seu teste de avaliação foi excedido!", "Teste de avaliação! ", (MessageBoxButtons)MessageBoxButton.OK, MessageBoxIcon.Error);
                        var = true;
                    }
                    else
                    {
                        TimeSpan remainingTime = TimeSpan.FromDays(30) - diff;
                        System.Windows.Forms.MessageBox.Show("Bem-vindo! Seu período de avaliação ainda está ativo." +
                                                              "\nTempo restante: \n" +
                                                              $"{remainingTime.Days} dias {remainingTime.Hours}':{remainingTime.Minutes}\"", "Inicialiação", (MessageBoxButtons)MessageBoxButton.OK, MessageBoxIcon.Information);
                    }
                }
            } */

            StartRibbonTab();
            System.Windows.Forms.MessageBox.Show("Para iniciar digite o comando:\n\nSTC-SHEET-TO-CAD\n\nOu acesse o menu Sheet To CAD", "Inicialiação", (MessageBoxButtons)MessageBoxButton.OK, MessageBoxIcon.Information);
        }


        List<Point3d> neutralBusPos = new List<Point3d>();

        Point3d mainBusStartPt;
        Point3d mainBusEndPt;
        Point3d mainBusMiddle;

        Point3d midPtBreakPos;

        Point3d firstPt;

        Point3d neutralSPDLocation = new Point3d();

        Point3d groundSPDPos = new Point3d();
        Point3d groundSPDMidPoint = new Point3d();

        Point3d lastPtLeft = new Point3d(); //ultimo ponto do lado esquerdo
        Point3d lastPtBottom = new Point3d();//ultimo ponto do fundo
        Point3d lastPtUp = new Point3d(); //ultimo ponto do topo
        Point3d lastPtRight = new Point3d(); //ultimo ponto do lado direito

        Point3d SPDWiringAlign = new Point3d();

        private const string _unipolar = "UNIPOLAR";
        private const string _bipolar = "BIPOLAR";
        private const string _tripolar = "TRIPOLAR";

        double firstLineLength = -225;//variaveis para desenhar linhas do diagrama
        double secondLineLengthBreaker = -190;
        double secondLineLengthWithDR = -69;
        double breakerLength = -36;

        double arrowMiddle = 3.375;
        double arrowLength = 8.25;

        double busSpace = 30;

        double DRLength = -63;
        double busConctDRlength = -29;

        const double cSXPos = -110; //conductor section x position
        const double mdTDisplacement = 8; //middle text (insertion point) displacement
        const double phaseLineXPos = -120;
        const double phaseTextXPos = -30;
        const double phaseLength = 15;

        const double neutralBusDisplac = -250;
        const double neutralBusLength = 40;
        const double neutralLineXpos = -310;

        double nextCircuit = -30;

        const int red = 1; //index de cores autocad
        const int cyan = 4;
        const int green = 3;
        const int white = 7;
        const int gray = 8;

        private string SPDCable = "";

        public List<string> RDCCables = new List<string>();

        private List<string> neutralDescGeneral = new List<string>();

        public string Voltage { get; set; } = "380 V";

        public string Source { get; set; } = "VEM DO QMP";
        public string Tube { get; set; } = "1.1/4\" - PVC";
        public string lineTypeBox { get; set; } = "HIDDEN";

        public string lineTypeCircuit { get; set; } = "Continuous";
        public int FontColor { get; set; } = white;
        public string FontName { get; set; } = "Arial";

        public double FontSize { get; set; } = 12.0;
        public AttachmentPoint AttPoint { get; set; } = AttachmentPoint.MiddleLeft;
        public int neutralColor { get; set; } = cyan;

        public double scaleFactor = 50.0;

        public string GroundingSystem { get; set; } = "Sistema de aterramento TN-S NBR-5410";

        public string BusSystem { get; set; } = "Pente 3F - 80 A";

        public string CircuitsLayer = "STC - Circuitos";
        public string NeutralLayer = "STC - Neutro";
        public string GroundLayer = "STC - Terra";
        public string BusLayer = "STC - Barramento";
        public string WiringLayer = "STC - Fiação";
        public string BoxLayer = "STC - Quadro";

        LayersNames LayersNames = new LayersNames();

        public string PhaseLaterA = "A";
        public string PhaseLaterB = "B";
        public string PhaseLaterC = "C";

        //pradrão de cabos para conexão no DR
        const float _RCD_25A = 4.0F;
        const float _RCD_40A = 6.0F;
        const float _RCD_63A = 10.0F;

        CultureInfo culture = new CultureInfo("fr-FR");
        RDCGroups rdcGroup = new RDCGroups();

        private enum Coluna
        {
            _default,
            Circuit, //coluna 1
            Description, //coluna 2
            Power, //coluna 3
            PhaseA, //coluna 6
            PhaseB, //coluna 7
            PhaseC, //coluna 8
            CircuitsCurrent, //coluna 9
            CablesQuantity, //coluna 10
            SplitCables, //coluna 11
            MinimizeEarth, //coluna 12
            ConductorSection, //coluna 13
            IsolationType, //coluna 14
            CBType, //coluna 15 ----    CB - Circuit Breaker
            CBCurve, //coluna 16
            CBPolesQuantity, //coluna 17
            CBCurrent, //coluna 18
            ShortCBCurrent, //coluna 19 -- corrente de curto circuito
            RCDNumbering, //coluna 20 ----        RCD - Residual Current Device
            RCDPolesQuant, //coluna 21
            RCDCurrent, //coluna 22
            RCDLeakage, //coluna 23
            Title
        };

        private bool ribbonTabCreated = false;

        [CommandMethod("StartSTC")]
        public void StartRibbonTab()
        {
            if (!ribbonTabCreated)
            {
                MenuTab ribbonMenu = new MenuTab();
                ribbonMenu.OpenMenu(this);
                ribbonTabCreated = true;
            }
        }

        static ProgressForm progressForm;

        public string totalTime = "";
        public string etapas = "";
        static void ProgressBar()
        {
            System.Windows.Forms.Application.Run(progressForm);
        }         
                

            [CommandMethod("STC-HELP")]
            public void openManual()
            {
                OpenManual openManual = new OpenManual();
                openManual.Open();
            }

            [CommandMethod("STC-SHEET-TO-CAD")]
            public void ShowInsetForm()
            {
                InsertForm istForm = new InsertForm(this);
                istForm.Show();
                if (istForm.DialogResult == System.Windows.Forms.DialogResult.OK)
                    istForm.Close();
            }


            public void ExcelToCAD(Point3d insPt, string pathFile, string sheetName)
            {
                //Thread th1 = new Thread(new ThreadStart(ProgressBar));
                //th1.Start();

                progressForm = new ProgressForm();
                progressForm.Show();
                        

                Stopwatch sw = new Stopwatch();
                sw.Start();

                firstPt = insPt;

                etapas = "Criando arquivos";
                progressForm.Invoke(new Action(() =>
                {
                    progressForm.UpdateProgress(10, etapas);
                }));

                CadManager cadManager = new CadManager();

                Document doc = cadManager.GetDocument();
                Database db = doc.Database;
                Editor edt = doc.Editor;

                doc.Editor.WriteMessage($"\n\nTempo de execução para criar arquivos: {sw.ElapsedMilliseconds} ms");
                etapas = "Lendo planilha";
                progressForm.Invoke(new Action(() =>
                {
                    progressForm.UpdateProgress(20, etapas);
                }));
                ExcelAccess LerExcel = new ExcelAccess();

                LerExcel.ReadColumn(pathFile, sheetName);

                doc.Editor.WriteMessage($"\n\nTempo de execução para ler planilha: {sw.ElapsedMilliseconds} ms");
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        doc.LockDocument();

                        BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                        BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        string relativeFolderPath = "config/acad.lin";

                        string currentDirectory = Environment.CurrentDirectory;

                        string folderPath = Path.Combine(currentDirectory, relativeFolderPath);

                        LoadLineType(doc, db, folderPath);

                        LayersNames.AddLayer(CircuitsLayer, white);
                        LayersNames.AddLayer(NeutralLayer, cyan);
                        LayersNames.AddLayer(GroundLayer, green);
                        LayersNames.AddLayer(BusLayer, red);
                        LayersNames.AddLayer(WiringLayer, white);
                        LayersNames.AddLayer(BoxLayer, white);

                        List<LayersNames> layerList = LayersNames.GetLayers();

                        CreateLayer(doc, db, layerList);

                        etapas = "Criando Layers";
                    progressForm.Invoke(new Action(() =>
                    {
                        progressForm.UpdateProgress(40, etapas);
                    }));

                        ReadRow(LerExcel, trans, btr, insPt);

                        etapas = "Criando desenho";
                    progressForm.Invoke(new Action(() =>
                    {
                        progressForm.UpdateProgress(70, etapas);
                    }));

                    trans.Commit();
                        doc.SendStringToExecute("._zoom e ", false, false, false);
                        doc.SendStringToExecute("._regen ", false, false, false);

                    doc.Editor.WriteMessage($"\n\nTempo de execução para executar desenho: {sw.ElapsedMilliseconds} ms");
                    etapas = "Finalizando";
                    progressForm.Invoke(new Action(() =>
                    {
                        progressForm.UpdateProgress(100, etapas);
                    }));

                    sw.Stop();
                    doc.Editor.WriteMessage($"\n\nTempo de execução total: {sw.ElapsedMilliseconds} ms");
                    totalTime = sw.ElapsedMilliseconds.ToString();
                    }
                    catch (System.Exception ex)
                    {
                        doc.Editor.WriteMessage("Ocorreu um erro: " + ex.Message);
                        trans.Abort();
                    }

                finally
                {
                    
                    progressForm.Invoke(new Action(() =>
                    {
                        progressForm.Close();
                    }));

                   // th1.Join();
                    
                }


                }

            }

            public void LoadLineType(Document doc, Database db, string folderPath)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        LinetypeTable ltTab = trans.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                        string ltName = "HIDDEN";
                        if (ltTab.Has(ltName))
                        {
                         //   doc.Editor.WriteMessage("Linetype already exist");
                            //trans.Abort();
                        }
                        else
                        {
                            db.LoadLineTypeFile(ltName, folderPath); //load the center linetype
                            doc.Editor.WriteMessage("Linetype [" + ltName + "] was created");
                            trans.Commit();
                        }
                        db.Ltscale = scaleFactor;
                    }
                    catch (System.Exception ex)
                    {
                        doc.Editor.WriteMessage("Error: " + ex.Message);
                        trans.Abort();
                    }
                }
            }

            public void CreateLayer(Document doc, Database db, List<LayersNames> layers)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    LayerTable lyTab = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                    foreach (LayersNames layer in layers)
                    {
                        if (lyTab.Has(layer.LayerName.ToString()))
                        {
                            doc.Editor.WriteMessage("Layer already exist.");
                            //trans.Abort();
                        }
                        else
                        {
                            lyTab.UpgradeOpen();
                            LayerTableRecord ltr = new LayerTableRecord();
                            ltr.Name = layer.LayerName.ToString();
                            ltr.Color = Color.FromColorIndex(ColorMethod.ByLayer, (short)layer.ColorIndex);
                            lyTab.Add(ltr);
                            trans.AddNewlyCreatedDBObject(ltr, true);
                            db.Clayer = lyTab[layer.LayerName.ToString()]; //define como a camada atual "Current Layer"
                        }
                    }

                    trans.Commit();
                }
            }

            private void DrawLinesCircuits(Transaction trans, BlockTableRecord btr, bool comunCircuit, ref Point3d lastPoint)
            {
                using (trans = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.TransactionManager.StartTransaction())
                {
                    try
                    {
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();

                        Line lineCircuitos = new Line(); //desenha linha de conexão com a seta

                        lineCircuitos.StartPoint = new Point3d((lastPoint.X), (lastPoint.Y + arrowMiddle), 0);
                        lineCircuitos.EndPoint = new Point3d((lastPoint.X + firstLineLength), (lastPoint.Y + arrowMiddle), 0);
                        lineCircuitos.Layer = CircuitsLayer;

                        btr.AppendEntity(lineCircuitos);
                        trans.AddNewlyCreatedDBObject(lineCircuitos, true);


                        Line lineCircuitos2 = new Line();

                        if (comunCircuit) //desenha linha de conexão com barramento
                        {
                            lineCircuitos2.StartPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthBreaker)), (lastPoint.Y + arrowMiddle), 0);
                            lineCircuitos2.EndPoint = new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0);
                            lineCircuitos2.Layer = CircuitsLayer;
                        }
                        else
                        {
                            lineCircuitos2.StartPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR)), (lastPoint.Y + arrowMiddle), 0);
                            lineCircuitos2.EndPoint = new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0);
                            lineCircuitos2.Layer = CircuitsLayer;
                        }
                        btr.AppendEntity(lineCircuitos2);
                        trans.AddNewlyCreatedDBObject(lineCircuitos2, true);

                        Arrow(trans, btr, new Point3d(lastPoint.X, lastPoint.Y + arrowMiddle, 0), LineWeight.ByLayer, lineTypeCircuit, CircuitsLayer);

                        trans.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Ocorreu um erro: ");
                        trans.Abort();
                    }
                }

            }

            private void DrawDescCircuits(ExcelAccess LerExcel, Transaction trans, BlockTableRecord btr, int rowCount, ref Point3d lastPoint) //desenha as linhas dos circuitos
            {
                string content = "";

                Description descricao = new Description();

                Point3d location = new Point3d();
                Point3d startPt;
                Point3d endPt;

                Wiring wiring = new Wiring();
                Bus faseLine = new Bus();

                ExcelAccess.Row row;
                row = LerExcel.rows[rowCount];

                if (row.CBPolesQuantity == _unipolar) InserirBloco("1xDIN.dwg", new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0));
                if (row.CBPolesQuantity == _bipolar) InserirBloco("2xDIN.dwg", new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0));
                if (row.CBPolesQuantity == _tripolar) InserirBloco("3xDIN.dwg", new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0));

                content = row.Circuit + " - " + row.Description;
                location = new Point3d((lastPoint.X + 15), (lastPoint.Y + arrowMiddle), 0); //distancia para inseção da descrição é de 15 da ponta da seta           
                descricao.AddDescription(location, FontColor, content, AttPoint, FontSize, 0, FontName);

                content = row.CBCurve + "-" + row.CBCurrent;
                location = new Point3d((lastPoint.X + firstLineLength), (lastPoint.Y + arrowMiddle + mdTDisplacement), 0);
                descricao.AddDescription(location, FontColor, content, AttPoint, FontSize, 0, FontName);

                content = row.ConductorSection;
                location = new Point3d((lastPoint.X + cSXPos), (lastPoint.Y + arrowMiddle + mdTDisplacement), 0);
                descricao.AddDescription(location, FontColor, content, AttPoint, FontSize, 0, FontName);


                if (row.CBPolesQuantity == _unipolar) //desenha fiação dos circuitos
                {
                    wiring.AddPhase(new Point3d((lastPoint.X + phaseLineXPos), (lastPoint.Y + arrowMiddle), 0), LineWeight.ByLayer, WiringLayer, "F");
                }
                else if (row.CBPolesQuantity == _bipolar)
                {
                    wiring.AddPhase(new Point3d((lastPoint.X + phaseLineXPos), (lastPoint.Y + arrowMiddle), 0), LineWeight.ByLayer, WiringLayer, "2F");
                }
                else
                {
                    wiring.AddPhase(new Point3d((lastPoint.X + phaseLineXPos), (lastPoint.Y + arrowMiddle), 0), LineWeight.ByLayer, WiringLayer, "3F");
                }

                if (!string.IsNullOrEmpty(row.PhaseA) && string.IsNullOrEmpty(row.PhaseB))
                {
                    DrawPhase(PhaseLaterA, white, lastPoint, arrowMiddle, mdTDisplacement);
                }

                if (!string.IsNullOrEmpty(row.PhaseB) && string.IsNullOrEmpty(row.PhaseA))
                {
                    DrawPhase(PhaseLaterB, gray, lastPoint, arrowMiddle, mdTDisplacement);
                }

                if (!string.IsNullOrEmpty(row.PhaseC) && string.IsNullOrEmpty(row.PhaseB) && string.IsNullOrEmpty(row.PhaseA))
                {
                    DrawPhase(PhaseLaterC, red, lastPoint, arrowMiddle, mdTDisplacement);
                }

                if (!string.IsNullOrEmpty(row.PhaseA) && !string.IsNullOrEmpty(row.PhaseB) && string.IsNullOrEmpty(row.PhaseC))
                {
                    DrawPhase(PhaseLaterA, white, lastPoint, arrowMiddle, mdTDisplacement);
                    DrawPhase(PhaseLaterB, gray, new Point3d(lastPoint.X + 12, lastPoint.Y, 0), arrowMiddle, mdTDisplacement);
                }
                if (!string.IsNullOrEmpty(row.PhaseA) && !string.IsNullOrEmpty(row.PhaseC) && string.IsNullOrEmpty(row.PhaseB))
                {
                    DrawPhase(PhaseLaterA, white, lastPoint, arrowMiddle, mdTDisplacement);
                    DrawPhase(PhaseLaterC, red, new Point3d(lastPoint.X + 12, lastPoint.Y, 0), arrowMiddle, mdTDisplacement);
                }
                if (!string.IsNullOrEmpty(row.PhaseB) && !string.IsNullOrEmpty(row.PhaseC) && string.IsNullOrEmpty(row.PhaseA))
                {
                    DrawPhase(PhaseLaterB, gray, lastPoint, arrowMiddle, mdTDisplacement);
                    DrawPhase(PhaseLaterC, red, new Point3d(lastPoint.X + 12, lastPoint.Y, 0), arrowMiddle, mdTDisplacement);
                }

                if (!string.IsNullOrEmpty(row.PhaseA) && !string.IsNullOrEmpty(row.PhaseB) && !string.IsNullOrEmpty(row.PhaseC))
                {
                    DrawTrifasic(lastPoint, arrowMiddle, mdTDisplacement);
                }

                // Função para desenhar uma fase
                void DrawPhase(string phase, int color, Point3d _lastPoint, double arrowMiddle, double mdTDisplacement)
                {
                    location = new Point3d((_lastPoint.X + phaseTextXPos), (_lastPoint.Y + arrowMiddle + mdTDisplacement), 0);
                    descricao.AddDescription(location, color, phase, AttachmentPoint.MiddleLeft, FontSize, 0, FontName);
                }

                // Função para desenhar as fases ABC (trifásico)
                void DrawTrifasic(Point3d _lastPoint, double arrowMiddle, double mdTDisplacement)
                {
                    DrawPhase(PhaseLaterA, white, new Point3d((_lastPoint.X - 43 - phaseTextXPos), (_lastPoint.Y), 0), arrowMiddle, mdTDisplacement);
                    DrawPhase(PhaseLaterB, gray, new Point3d((_lastPoint.X - 30 - phaseTextXPos), (_lastPoint.Y), 0), arrowMiddle, mdTDisplacement);
                    DrawPhase(PhaseLaterC, red, new Point3d((_lastPoint.X - 18 - phaseTextXPos), (_lastPoint.Y), 0), arrowMiddle, mdTDisplacement);
                }

                descricao.DrawDescription(trans, btr);
                faseLine.DrawBus(trans, btr);
                wiring.DrawWires(trans, btr);

                lastPoint = new Point3d(lastPoint.X, lastPoint.Y + nextCircuit, 0); //Acrescenta o deslocamento para o próximo circuito

            }

            private void NeutralBusRDC(ExcelAccess LerExcel, Transaction trans, BlockTableRecord btr, int count, int rowCount, ref Point3d lastPoint) //
            {
                string content = "";
                Description descricao = new Description();
                Wiring wiring = new Wiring();

                Point3d location = new Point3d();
                Point3d startPt;
                Point3d endPt;

                Bus neutralLine = new Bus();

                ExcelAccess.Row row;
                row = LerExcel.rows[rowCount - 1];

                content = "(X" + count + ")"; //insere descrições do barramento do neutro
                location = new Point3d((lastPoint.X + cSXPos), (lastPoint.Y + arrowMiddle + mdTDisplacement), 0);
                descricao.AddDescription(location, white, content, AttachmentPoint.MiddleLeft, 10.0, 0, FontName);

                content = "BARRA NEUTRO" + "\nIDR-" + row.RCDNumbering;
                location = new Point3d((lastPoint.X + firstLineLength + 50), (lastPoint.Y + arrowMiddle + mdTDisplacement + 5), 0);
                descricao.AddDescription(location, white, content, AttachmentPoint.MiddleCenter, 10.0, 0, FontName);

                startPt = new Point3d((lastPoint.X + neutralBusDisplac), (lastPoint.Y + neutralBusLength / 2), 0); //desenha barramento
                endPt = new Point3d((lastPoint.X + neutralBusDisplac), (lastPoint.Y - neutralBusLength / 2), 0);
                neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                wiring.AddPhase(new Point3d((lastPoint.X + neutralLineXpos), (lastPoint.Y), 0), LineWeight.ByLayer, WiringLayer, "N"); //desenha fiação do neutro p/ circuitos do DR

                location = new Point3d((lastPoint.X + 15), (lastPoint.Y), 0);
                content = "NEUTROS - CIRCUITOS " + (rowCount - count + 1) + " A " + rowCount;
                descricao.AddDescription(location, white, content, AttachmentPoint.MiddleLeft, FontSize, 0, FontName);

                switch (row.RCDCurrent)
                {
                    case "25 A":
                        content = "#" + _RCD_25A.ToString("F1", culture);
                        break;
                    case "40 A":
                        content = "#" + _RCD_40A.ToString("F1", culture);
                        break;
                    case "63 A":
                        content = "#" + _RCD_63A.ToString("F1", culture);
                        break;
                    default:
                        content = "";
                        break;
                }
                location = new Point3d((lastPoint.X + neutralLineXpos + 5), (lastPoint.Y + arrowMiddle + mdTDisplacement), 0);
                descricao.AddDescription(location, white, content, AttachmentPoint.MiddleLeft, 10.0, 0, FontName);


                descricao.DrawDescription(trans, btr);
                neutralLine.DrawBus(trans, btr);
                wiring.DrawWires(trans, btr);
            }

            private void DrawBus(ExcelAccess LerExcel, Transaction trans, BlockTableRecord btr, bool comunCircuit, int count, int rowCount, ref Point3d lastPoint)
            {
                Point3d startPoint;
                Point3d endPoint;

                Point3d posNeutroCon;

                Description description = new Description();

                Bus barramento = new Bus();

                ExcelAccess.Row row = LerExcel.rows[rowCount - count];

                if (comunCircuit) //Desenha barramento geral
                {
                    startPoint = new Point3d((firstPt.X + firstLineLength + (secondLineLengthBreaker)), (lastPoint.Y + arrowMiddle), 0);
                    endPoint = new Point3d((firstPt.X + firstLineLength + (secondLineLengthBreaker)), (firstPt.Y + arrowMiddle + busSpace), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.LineWeight080, "Continuous", BusLayer);

                    mainBusStartPt = startPoint;
                    mainBusEndPt = endPoint;
                    mainBusMiddle = new Point3d((firstPt.X + firstLineLength + (secondLineLengthBreaker)), (startPoint.Y + endPoint.Y) / 2, 0);
                    //adiciona descrição do barramento geral
                    description.AddDescription(new Point3d(startPoint.X - 3, startPoint.Y, 0), FontColor, BusSystem, AttachmentPoint.BottomLeft, FontSize, 1.5707963, FontName);

                    neutralBus(LerExcel, trans, btr);
                }
                if (!comunCircuit && count > 1) //Desenha barramento dos DRs se tiver mais de um circuito no grupo
                {
                    //barramento vertical:
                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR)), (lastPoint.Y + arrowMiddle), 0);
                    endPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR)), (lastPoint.Y + arrowMiddle + (count * busSpace + busSpace)), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.LineWeight080, "Continuous", BusLayer);

                    //barramentos de conexão do DR ao barramento princial e ao conjunto de disjuntores
                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR)), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace)), 0);
                    endPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace)), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", CircuitsLayer);
                    //fase
                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength + DRLength), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace)), 0);
                    endPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + 2 * busConctDRlength + DRLength), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace)), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", CircuitsLayer);


                    //neutro sub-barramento para o DR
                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength - 2.3463), (lastPoint.Y + arrowMiddle - 2.3277 + (((double)count + 1) / 2 * busSpace) - 22.5), 0);
                    endPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength - 2.3463), (lastPoint.Y + arrowMiddle - ((((double)count + 1) / 2) * busSpace) - 22.5), 0);
                    Point3d midPoint = new Point3d((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2 - 6.3362, 0);
                    barramento.AddBus(startPoint, midPoint, LineWeight.ByLayer, "Continuous", NeutralLayer);

                    //desenha linha horizontal do neutro para ligar a seta na sequencia
                    endPoint = new Point3d(midPoint.X + 325.3463, midPoint.Y, 0);
                    barramento.AddBus(midPoint, endPoint, LineWeight.ByLayer, "Continuous", NeutralLayer);

                    NeutralBusRDC(LerExcel, trans, btr, count, rowCount, ref endPoint); //adiciona descrição barramento neutro

                    Arrow(trans, btr, new Point3d(endPoint.X, (endPoint.Y), 0), LineWeight.ByLayer, lineTypeCircuit, NeutralLayer);//desenha seta do neutro


                    //neutro conexão com barramento geral 
                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength + DRLength), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace) - 22.5), 0);
                    endPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + 2 * busConctDRlength + DRLength + 5), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace) - 22.5), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", NeutralLayer);

                    posNeutroCon = endPoint;

                    Autodesk.AutoCAD.DatabaseServices.Polyline neutralCurve = new Autodesk.AutoCAD.DatabaseServices.Polyline(); //desenha curva de desvio barramento geral 
                    neutralCurve.AddVertexAt(0, new Point2d(posNeutroCon.X, posNeutroCon.Y), 0, 0, 0);
                    neutralCurve.AddVertexAt(0, new Point2d(posNeutroCon.X, posNeutroCon.Y + 5), 0, 0, 0);
                    neutralCurve.AddVertexAt(0, new Point2d(posNeutroCon.X - 10, posNeutroCon.Y + 5), 0, 0, 0);
                    neutralCurve.AddVertexAt(0, new Point2d(posNeutroCon.X - 10, posNeutroCon.Y), 0, 0, 0);
                    neutralCurve.Layer = NeutralLayer;
                    btr.AppendEntity(neutralCurve);
                    trans.AddNewlyCreatedDBObject(neutralCurve, true);

                    neutralBusPos.Add(new Point3d(posNeutroCon.X - 10, posNeutroCon.Y, 0)); //adiciona na lista as posições de fiação para conectar o barramento de neutro posteriormente


                }

                if (!comunCircuit && count == 1) //barramentos para dr com apenas um circuito
                {
                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR)), (lastPoint.Y + arrowMiddle + (busSpace)), 0);
                    endPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength), (lastPoint.Y + arrowMiddle + (busSpace)), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", CircuitsLayer);

                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength + DRLength), (lastPoint.Y + arrowMiddle + (busSpace)), 0);
                    endPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + 2 * busConctDRlength + DRLength), (lastPoint.Y + arrowMiddle + (busSpace)), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", CircuitsLayer);

                    //seta barramento
                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace) - 22.5), 0);
                    endPoint = new Point3d(lastPoint.X, (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace) - 22.5), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", NeutralLayer);

                    Arrow(trans, btr, new Point3d(endPoint.X, (endPoint.Y), 0), LineWeight.ByLayer, lineTypeCircuit, NeutralLayer);//desenha seta do neutro   

                    //neutro conexão com barramento geral
                    startPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength + DRLength), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace) - 22.5), 0);
                    endPoint = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + 2 * busConctDRlength + DRLength + 5), (lastPoint.Y + arrowMiddle + (((double)count + 1) / 2 * busSpace) - 22.5), 0);
                    barramento.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", NeutralLayer);

                    posNeutroCon = endPoint;

                    Autodesk.AutoCAD.DatabaseServices.Polyline neutralCurve = new Autodesk.AutoCAD.DatabaseServices.Polyline(); //desenha curva de desvio barramento geral 
                    neutralCurve.AddVertexAt(0, new Point2d(posNeutroCon.X, posNeutroCon.Y), 0, 0, 0);
                    neutralCurve.AddVertexAt(0, new Point2d(posNeutroCon.X, posNeutroCon.Y + 5), 0, 0, 0);
                    neutralCurve.AddVertexAt(0, new Point2d(posNeutroCon.X - 10, posNeutroCon.Y + 5), 0, 0, 0);
                    neutralCurve.AddVertexAt(0, new Point2d(posNeutroCon.X - 10, posNeutroCon.Y), 0, 0, 0);
                    neutralCurve.Layer = NeutralLayer;
                    btr.AppendEntity(neutralCurve);
                    trans.AddNewlyCreatedDBObject(neutralCurve, true);

                    neutralBusPos.Add(new Point3d(posNeutroCon.X - 10, posNeutroCon.Y, 0)); //adiciona na lista as posições de fiação para conectar o barramento de neutro posteriormente


                }

                description.DrawDescription(trans, btr);
                barramento.DrawBus(trans, btr);

            }

            public void neutralBus(ExcelAccess rdExcel, Transaction trans, BlockTableRecord btr)
            {

                Bus neutralLine = new Bus();

                Wiring wiring = new Wiring();

                string content = "";
                Description descricao = new Description();
                Point3d location = new Point3d();

                double deslocPoint = 10;
                double sapacing = 0;

                Point3d startPt;
                Point3d endPt;

                Point3d neutralBuslocation;

                ExcelAccess.Row row;
                row = rdExcel.rows[0];

                try
                {

                    for (int i = 0; i < rdExcel.others.Count; i++)
                    {
                        if (rdExcel.others[i].Circuit == "DPS")
                        {
                            SPDCable = rdExcel.others[i].ConductorSection;
                        }
                    }

                    startPt = new Point3d(mainBusEndPt.X - neutralBusPos.Count * 15 - 100, mainBusEndPt.Y + (neutralBusPos.Count * 20 / 2) + 30, 0); //desenha linha do barramento geral
                    endPt = new Point3d(mainBusEndPt.X - neutralBusPos.Count * 15 - 100, mainBusEndPt.Y - (neutralBusPos.Count * 20 / 2) - 30, 0);

                    content = "BARRA NEUTRO \n GERAL";
                    location = new Point3d((startPt.X + 3), (startPt.Y), 0);
                    descricao.AddDescription(location, FontColor, content, AttachmentPoint.MiddleLeft, 10.0, 0, FontName);

                    lastPtUp = startPt; //pega a ultima linha desenha no topo

                    neutralLine.AddBus(startPt, endPt, LineWeight.LineWeight080, "Continuous", NeutralLayer);

                    neutralBuslocation = new Point3d(startPt.X, startPt.Y, 0);

                    startPt = new Point3d(startPt.X, startPt.Y - 30, 0);
                    endPt = new Point3d(firstPt.X, startPt.Y, 0);

                    //desenha seta geral barramento do neutro
                    neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                    Arrow(trans, btr, endPt, LineWeight.ByLayer, lineTypeCircuit, NeutralLayer); //desenha seta do neutro

                    wiring.AddPhase(new Point3d(endPt.X + phaseLineXPos, endPt.Y, 0), LineWeight.ByLayer, WiringLayer, "N"); //desenha fiação do neutro

                    content = "NEUTRO - CIRCUITOS ";//descrição geral dos circuitos
                    location = new Point3d((endPt.X + 15), (endPt.Y), 0);
                    descricao.AddDescription(location, FontColor, content, AttPoint, FontSize, 0, FontName);


                    //desenha conexão do barramento de neutro na alimentação de entrada
                    startPt = new Point3d(neutralBuslocation.X, neutralBuslocation.Y - (neutralBusPos.Count * 20 / 2) - 30, 0);
                    endPt = new Point3d(startPt.X - 30, startPt.Y, 0);
                    neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);


                    startPt = endPt;
                    endPt = new Point3d(startPt.X, mainBusMiddle.Y + 15, 0);
                    neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                    wiring.AddPhase(new Point3d(endPt.X, (startPt.Y + endPt.Y) / 2, 0), LineWeight.ByLayer, WiringLayer, "N Bus"); //desenha fiação do neutro no barramento principal

                    for (int j = 0; j < rdExcel.others.Count; j++)
                    {
                        if (rdExcel.others[j].Circuit == "ENTRADA DE ENERGIA")
                        {
                            content = rdExcel.others[j].ConductorSection;
                            location = new Point3d(endPt.X, 5 + (startPt.Y + endPt.Y) / 2, 0);
                            descricao.AddDescription(location, white, content, AttachmentPoint.BottomLeft, FontSize, 0, FontName);
                        }
                    }

                    startPt = endPt;
                    endPt = new Point3d(startPt.X - 120, startPt.Y, 0);
                    neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                    lastPtLeft = endPt;

                    startPt = new Point3d(mainBusMiddle.X - (neutralBusPos.Count - 1) * 10 - 15, mainBusMiddle.Y, 0); //desenha linha do barramento geral

                    midPtBreakPos = new Point3d(startPt.X + (-(neutralBusPos.Count * 15 + 100 - ((neutralBusPos.Count - 1) * (10) + 15)) / 2 - 15 + (breakerLength / 2)), mainBusMiddle.Y, 0); //essa posição deu trabalho pqp

                    /*atenção verificar bug a partir daqui
                    *
                    *
                    *
                    *
                    */
                    for (int i = 0; i < neutralBusPos.Count; i++) //linhas de ligação aos drs
                    {
                        if ((i + 1) == neutralBusPos.Count) //desenha o neutro do dps
                        {
                            startPt = neutralSPDLocation;
                            endPt = new Point3d(neutralSPDLocation.X - 77, neutralSPDLocation.Y, 0);
                            neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                            SPDWiringAlign = new Point3d((startPt.X + endPt.X) / 2, startPt.Y, 0);

                            wiring.AddPhase(new Point3d(SPDWiringAlign.X, SPDWiringAlign.Y, 0), LineWeight.ByLayer, WiringLayer, "N"); //desenha fiação do neutro no dps

                            content = SPDCable;
                            location = new Point3d(SPDWiringAlign.X, SPDWiringAlign.Y, 0);
                            descricao.AddDescription(location, white, content, AttachmentPoint.BottomLeft, FontSize, 0, FontName);

                            startPt = endPt;
                            endPt = new Point3d(startPt.X, startPt.Y - 45, 0);
                            neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                            startPt = endPt;
                            endPt = new Point3d(startPt.X - 77 - (neutralBusPos.Count - 1) * 10 - 15, startPt.Y, 0);
                            neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                            neutralSPDLocation = endPt;

                            startPt = endPt;
                            endPt = new Point3d(endPt.X, neutralBuslocation.Y - 50 - sapacing, 0); //sobe ao barramento
                            sapacing += 20;
                            neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                            startPt = endPt;
                            endPt = new Point3d(neutralBuslocation.X, endPt.Y, 0); //linha vertical de conexão no barramento
                            neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                            wiring.AddPhase(new Point3d(endPt.X + 20, endPt.Y, 0), LineWeight.ByLayer, WiringLayer, "N"); //desenha fiação do neutro no barramento geral

                            content = SPDCable; //descrição cabeamento dos DRS
                            location = new Point3d(endPt.X + 20, endPt.Y, 0);
                            descricao.AddDescription(location, white, content, AttachmentPoint.BottomLeft, FontSize, 0, FontName);


                        }

                        if ((i + 1) < neutralBusPos.Count) //desenha o neutro dos DRs
                        {
                            startPt = neutralBusPos[i];
                            endPt = new Point3d(startPt.X - deslocPoint, startPt.Y, 0); //linha horizontal
                            neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);
                            deslocPoint += 10;

                            startPt = endPt;
                            endPt = new Point3d(endPt.X, neutralBuslocation.Y - 50 - sapacing, 0); //sobe ao barramento
                            sapacing += 20;
                            neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                            startPt = endPt;
                            endPt = new Point3d(neutralBuslocation.X, endPt.Y, 0); //linha vertical de conexão no barramento
                            neutralLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", NeutralLayer);

                            wiring.AddPhase(new Point3d(endPt.X + 20, endPt.Y, 0), LineWeight.ByLayer, WiringLayer, "N"); //desenha fiação do neutro

                            content = RDCCables[i]; //descrição cabeamento dos DRS
                            location = new Point3d(endPt.X + 20, endPt.Y, 0);
                            descricao.AddDescription(location, white, content, AttachmentPoint.BottomLeft, FontSize, 0, FontName);
                        }
                    }

                    RDCCables.Clear();

                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Ocorreu um erro: " + ex.Message, "Erro!", (MessageBoxButtons)MessageBoxButton.OK, MessageBoxIcon.Error);
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Ocorreu um erro: ");
                    trans.Abort();
                }


                descricao.DrawDescription(trans, btr);
                neutralLine.DrawBus(trans, btr);
                wiring.DrawWires(trans, btr);

                GroundBus(rdExcel, trans, btr);
            }


            //verificar bug aqui
            // ***
            // ***
            // ***
            public void GroundBus(ExcelAccess rdExcel, Transaction trans, BlockTableRecord btr)
            {
                Bus groundLine = new Bus();

                Wiring wiring = new Wiring();

                string content = "";
                Description descricao = new Description();
                Point3d location = new Point3d();

                Point3d startPt;
                Point3d endPt;

                Point3d mainGroundBusPt;

                startPt = new Point3d(mainBusStartPt.X - neutralBusPos.Count * 15 - 100, mainBusStartPt.Y + 45, 0); //desenha linha do barramento geral
                endPt = new Point3d(mainBusStartPt.X - neutralBusPos.Count * 15 - 100, mainBusStartPt.Y - 45, 0);
                mainGroundBusPt = startPt;

                content = "BARRA TERRA \n GERAL";
                location = new Point3d((startPt.X + 3), (startPt.Y), 0);
                descricao.AddDescription(location, FontColor, content, AttachmentPoint.MiddleLeft, 10.0, 0, FontName);
                groundLine.AddBus(startPt, endPt, LineWeight.LineWeight080, "Continuous", GroundLayer);

                startPt = new Point3d(startPt.X, startPt.Y - 30, 0); //fiação que conecta ao DPS (terra)
                endPt = new Point3d(startPt.X + 100, startPt.Y, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                startPt = endPt;
                endPt = new Point3d(startPt.X, mainBusStartPt.Y - 30, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                wiring.AddPhase(new Point3d(SPDWiringAlign.X, endPt.Y, 0), LineWeight.ByLayer, WiringLayer, "T"); //desenha fiação

                content = SPDCable;
                location = new Point3d(SPDWiringAlign.X + 5, endPt.Y, 0);
                descricao.AddDescription(location, white, content, AttachmentPoint.BottomLeft, FontSize, 0, FontName);

                startPt = endPt;
                endPt = new Point3d(groundSPDPos.X + 30, startPt.Y, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                startPt = endPt;
                endPt = new Point3d(startPt.X, groundSPDMidPoint.Y, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                startPt = endPt;
                endPt = groundSPDMidPoint;
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);


                //desenha condutor geral de aterramento dos circuitos 
                startPt = new Point3d(mainGroundBusPt.X, mainGroundBusPt.Y - 60, 0);
                endPt = new Point3d(startPt.X + 85, startPt.Y, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                startPt = endPt;
                endPt = new Point3d(startPt.X, startPt.Y - 30, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                startPt = endPt;
                endPt = new Point3d(firstPt.X, startPt.Y, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                content = "TERRA - CIRCUITOS "; //descrição geral dos circuitos
                location = new Point3d((endPt.X + 15), (endPt.Y), 0);
                descricao.AddDescription(location, FontColor, content, AttPoint, FontSize, 0, FontName);

                Arrow(trans, btr, endPt, LineWeight.ByLayer, lineTypeCircuit, GroundLayer); //desenha seta do terra

                wiring.AddPhase(new Point3d(endPt.X + phaseLineXPos, endPt.Y, 0), LineWeight.ByLayer, WiringLayer, "T");

                lastPtBottom = endPt; //pega o ultimo ponto desenhado no diagrama da parte de baixo


                //desenha conexão do barramento com a alimentação de entrada
                startPt = new Point3d(mainGroundBusPt.X, mainGroundBusPt.Y - 45, 0);
                endPt = new Point3d(startPt.X - 30, startPt.Y, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                startPt = endPt;
                endPt = new Point3d(startPt.X, mainBusMiddle.Y - 15, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);

                wiring.AddPhase(new Point3d(endPt.X, (startPt.Y + endPt.Y) / 2, 0), LineWeight.ByLayer, WiringLayer, "T Bus"); //desenha fiação do terra ao barramento principal

                for (int j = 0; j < rdExcel.others.Count; j++)
                {
                    if (rdExcel.others[j].Circuit == "ENTRADA DE ENERGIA")
                    {
                        content = rdExcel.others[j].ConductorSection;
                        location = new Point3d(endPt.X, 5 + (startPt.Y + endPt.Y) / 2, 0);
                        descricao.AddDescription(location, white, content, AttachmentPoint.BottomLeft, FontSize, 0, FontName);
                    }
                }

                startPt = endPt;
                endPt = new Point3d(startPt.X - 120, startPt.Y, 0);
                groundLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", GroundLayer);


                descricao.DrawDescription(trans, btr);
                groundLine.DrawBus(trans, btr);
                wiring.DrawWires(trans, btr);
            }

            public void InsertMainBreaker(ExcelAccess rdExcel, Transaction trans, BlockTableRecord btr, int rowCount, ref Point3d lastPoint)
            {
                string content = "";
                Description descricao = new Description();
                Point3d location = new Point3d();
                Point3d startPt;
                Point3d endPt;

                double majorRadious = 20.0;
                double minorRadious = 3;

                Bus busLine = new Bus();

                Wiring wiring = new Wiring();

                ExcelAccess.Row breaker;
                breaker = rdExcel.others[rowCount];

                startPt = mainBusMiddle;
                endPt = new Point3d(midPtBreakPos.X - breakerLength, startPt.Y, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", CircuitsLayer);

                startPt = new Point3d(endPt.X + breakerLength, endPt.Y, 0);
                endPt = new Point3d(lastPtLeft.X, startPt.Y, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", CircuitsLayer);

                if (breaker.CBPolesQuantity == _unipolar) InserirBloco("1xDIN.dwg", midPtBreakPos);
                if (breaker.CBPolesQuantity == _bipolar) InserirBloco("2xDIN.dwg", midPtBreakPos);
                if (breaker.CBPolesQuantity == _tripolar) InserirBloco("3xDIN.dwg", midPtBreakPos);

                //insere dados do disjuntor geral 
                content = breaker.Description + "\n" + breaker.CBCurve + "-" + breaker.CBCurrent;
                location = new Point3d((midPtBreakPos.X - breakerLength / 2), (midPtBreakPos.Y + 25), 0);
                descricao.AddDescription(location, FontColor, content, AttachmentPoint.BottomCenter, FontSize, 0, FontName);

                content = Voltage + "/" + breaker.ShortCBCurrent;
                location = new Point3d((midPtBreakPos.X - breakerLength / 2), (midPtBreakPos.Y - 8), 0);
                descricao.AddDescription(location, FontColor, content, AttachmentPoint.TopCenter, FontSize, 0, FontName);

                //insere dados da alimentação de entrada
                content = Source;
                location = new Point3d((lastPtLeft.X - 10), (startPt.Y), 0);
                descricao.AddDescription(location, FontColor, content, AttachmentPoint.MiddleRight, FontSize, 0, "Arial Black");

                //insere dados da fiação da alimentação de entrada e tubulação

                location = new Point3d((lastPtLeft.X + 60), (startPt.Y), 0);
                Ellipse ellipseCableIndicator = new Ellipse(location, Vector3d.ZAxis, new Vector3d(0, -majorRadious, 0), minorRadious / majorRadious, 0, 0);
                ellipseCableIndicator.Layer = WiringLayer;
                btr.AppendEntity(ellipseCableIndicator);
                trans.AddNewlyCreatedDBObject(ellipseCableIndicator, true);

                startPt = new Point3d(location.X, location.Y - majorRadious, 0);
                endPt = new Point3d(location.X, startPt.Y - 20, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", WiringLayer);

                startPt = endPt;
                endPt = new Point3d(startPt.X - 60, startPt.Y, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", WiringLayer);

                string pipeE = "Ø" + Tube + "-PVC";

                if (breaker.CablesQuantity == "F+N+T")
                {
                    content = "3" + breaker.ConductorSection + "mm²" + "(" + breaker.IsolationType + ")"
                              + "\n" + pipeE;
                    wiring.AddPhase(new Point3d(endPt.X + 25, endPt.Y, 0), LineWeight.ByLayer, WiringLayer, breaker.CablesQuantity.ToString()); //desenha fiação

                }
                else if (breaker.CablesQuantity == "2F+N+T")
                {
                    content = "3" + breaker.ConductorSection + "mm²" + "(" + breaker.ConductorSection + ")" + "(" + breaker.IsolationType + ")"
                              + "\n" + pipeE;
                    wiring.AddPhase(new Point3d(endPt.X + 22.5, endPt.Y, 0), LineWeight.ByLayer, WiringLayer, breaker.CablesQuantity.ToString()); //desenha fiação
                }
                else
                {
                    content = "4" + breaker.ConductorSection + "mm²" + "(" + breaker.ConductorSection + ")" + "(" + breaker.IsolationType + ")"
                              + "\n" + pipeE;
                    wiring.AddPhase(new Point3d(endPt.X + 20, endPt.Y, 0), LineWeight.ByLayer, WiringLayer, breaker.CablesQuantity.ToString()); //desenha fiação
                }
                location = new Point3d((endPt.X + 30), (endPt.Y - 15), 0);
                descricao.AddDescription(location, FontColor, content, AttachmentPoint.TopCenter, 10.0, 0, FontName);


                wiring.DrawWires(trans, btr);
                descricao.DrawDescription(trans, btr);
                busLine.DrawBus(trans, btr);

            }

            public void DrawBox(ExcelAccess rdExcel, Transaction trans, BlockTableRecord btr, int rowCount, ref Point3d lastPoint)
            {
                Bus busLine = new Bus();
                Description description = new Description();

                ExcelAccess.Row row;
                row = rdExcel.rows[1];

                string content = "";

                Point3d startPt;
                Point3d endPt;

                Point3d location;

                content = row.Title;


                lastPtRight = firstPt;

                //draw a box around using the variables that indicate the bottons of the box (lastPtUp, lastPtRight, lastPtBottom, lastPtLeft), use a offset of 30
                startPt = new Point3d(lastPtLeft.X + 90, lastPtUp.Y + 30, 0); //desenha linha da esquerda
                endPt = new Point3d(lastPtLeft.X + 90, lastPtBottom.Y - 30, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, lineTypeBox, BoxLayer);

                endPt = new Point3d(lastPtRight.X - 60, startPt.Y, 0); //desenha linha do topo
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", BoxLayer);

                startPt = new Point3d(endPt.X, endPt.Y, 0); //desenha linha da direita
                endPt = new Point3d(startPt.X, lastPtBottom.Y - 30, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, lineTypeBox, BoxLayer);

                startPt = endPt;
                endPt = new Point3d(lastPtLeft.X + 90, endPt.Y, 0); //desenha linha de baixo
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, lineTypeBox, BoxLayer);
                InserirBloco("Grounding.dwg", startPt); //insere indicação de aterramento na caracaça

                //adiciona descrição do sistema de aterramento
                description.AddDescription(new Point3d(endPt.X, endPt.Y - 5, 0), FontColor, GroundingSystem, AttachmentPoint.TopLeft, FontSize, 0, FontName);

                //desenha o retangulo do titulo
                startPt = new Point3d(lastPtLeft.X + 90, lastPtUp.Y + 30, 0);
                endPt = new Point3d(startPt.X, startPt.Y + 30, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", BoxLayer);

                startPt = endPt;
                endPt = new Point3d(lastPtRight.X - 60, startPt.Y, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", BoxLayer);

                startPt = endPt;
                endPt = new Point3d(startPt.X, startPt.Y - 30, 0);
                busLine.AddBus(startPt, endPt, LineWeight.ByLayer, "Continuous", BoxLayer);


                //calculate the midpoint using the last points left and right as the center of the box and the y axis use the distance of 15 for the middle
                location = new Point3d((lastPtLeft.X + 90 + lastPtRight.X - 60) / 2, lastPtUp.Y + 30 + 15, 0);

                description.AddDescription(location, white, content, AttachmentPoint.MiddleCenter, 15.0, 0, "Arial Black");
                description.DrawDescription(trans, btr);

                busLine.DrawBus(trans, btr);

            }

            public void InsertDR(ExcelAccess rdExcel, Transaction trans, BlockTableRecord btr, int count, int rowCount, string idrPolos, ref Point3d lastPoint)
            {
                string content = "";
                string poles = "";
                double majorRadious = 15.0;
                double minorRadious = 2.1;

                Description description = new Description();
                Wiring wiring = new Wiring();
                Bus lines = new Bus();

                Point3d location = new Point3d();
                Point3d startPoint = new Point3d();
                Point3d endPoint = new Point3d();

                ExcelAccess.Row row;
                row = rdExcel.rows[rowCount - 1];

                Point3d lastPt = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength + DRLength),
                                (lastPoint.Y + arrowMiddle + ((double)count / 2 * busSpace + busSpace / 2)), 0);

                if (idrPolos == "BIPOLAR")
                {
                    InserirBloco("DR-2P.dwg", lastPt);
                    poles = "2P";
                }

                if (idrPolos == "TETRAPOLAR")
                {
                    InserirBloco("DR-4P.dwg", lastPt);
                    poles = "4P";
                }


                content = "DR-" + row.RCDNumbering + "\n" + poles + "-" + row.RCDCurrent + "\nFuga " + row.RCDLeakage;
                location = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength + DRLength / 2),
                                (lastPoint.Y + arrowMiddle + ((double)count / 2 * busSpace + busSpace / 2) + 15), 0); ///// atençãooooooooooo mudar 15 para espaçamento da inserção do texto 
                description.AddDescription(location, FontColor, content, AttachmentPoint.BottomCenter, FontSize, 0, FontName);
                description.DrawDescription(trans, btr);

                location = new Point3d((lastPoint.X + firstLineLength + (secondLineLengthWithDR) + busConctDRlength + DRLength + busConctDRlength / 2),
                                (lastPoint.Y + arrowMiddle + ((double)count / 2 * busSpace + busSpace / 2)) - 11.25, 0);

                Ellipse ellipseCableIndicator = new Ellipse(location, Vector3d.ZAxis, new Vector3d(0, -majorRadious, 0), minorRadious / majorRadious, 0, 0);
                ellipseCableIndicator.Layer = WiringLayer;
                btr.AppendEntity(ellipseCableIndicator);
                trans.AddNewlyCreatedDBObject(ellipseCableIndicator, true);


                startPoint = ellipseCableIndicator.StartPoint;
                endPoint = new Point3d(ellipseCableIndicator.StartPoint.X, ellipseCableIndicator.StartPoint.Y - 32, 0); //32 comprimento da linha de ligação da elipse

                lines.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", WiringLayer);

                startPoint = endPoint;
                endPoint = new Point3d(endPoint.X + 25, endPoint.Y, 0); // 25 comprimento da linha base para indicação da fiação do dr                       

                lines.AddBus(startPoint, endPoint, LineWeight.ByLayer, "Continuous", WiringLayer);

                if (row.RCDPolesQuant == "TETRAPOLAR")
                {
                    wiring.AddPhase(new Point3d(startPoint.X + 8, startPoint.Y, 0), LineWeight.ByLayer, WiringLayer, "3F+N");

                    switch (row.RCDCurrent)
                    {
                        case "25 A":
                            content = "4#" + _RCD_25A.ToString("F0", culture);
                            RDCCables.Add(content.Substring(1));
                            break;
                        case "40 A":
                            content = "4#" + _RCD_40A.ToString("F0", culture);
                            RDCCables.Add(content.Substring(1));
                            break;
                        case "63 A":
                            content = "4#" + _RCD_63A.ToString("F0", culture);
                            RDCCables.Add(content.Substring(1));
                            break;
                        default:
                            content = "";
                            break;
                    }
                }

                else
                {
                    wiring.AddPhase(new Point3d(startPoint.X + 10, startPoint.Y, 0), LineWeight.ByLayer, WiringLayer, "F+N");

                    switch (row.RCDCurrent)
                    {
                        case "25 A":
                            content = "2#" + _RCD_25A.ToString("F0", culture);
                            RDCCables.Add(content.Substring(1));
                            break;
                        case "40 A":
                            content = "2#" + _RCD_40A.ToString("F0", culture);
                            RDCCables.Add(content.Substring(1));
                            break;
                        case "63 A":
                            content = "2#" + _RCD_63A.ToString("F0", culture);
                            RDCCables.Add(content.Substring(1));
                            break;
                        default:
                            content = "";
                            break;
                    }
                }

                location = new Point3d(startPoint.X, startPoint.Y - 25, 0);
                description.AddDescription(location, white, content, AttachmentPoint.BottomLeft, 12.0, 0, "Arial");
                description.DrawDescription(trans, btr);
                lines.DrawBus(trans, btr);
                wiring.DrawWires(trans, btr);
            }

            private void InsertSPD(ExcelAccess rdExcel, Transaction trans, BlockTableRecord btr, int rowCount, ref Point3d lastPoint) //desenha as linhas dos circuitos
            {

                try
                {


                    string content = "";
                    Description descricao = new Description();
                    Point3d location = new Point3d();
                    Point3d startPt;
                    Point3d endPt;

                    Bus busLine = new Bus();
                    Wiring wiring = new Wiring();

                    ExcelAccess.Row SPD;
                    SPD = rdExcel.others[rowCount];

                    if (SPD.CablesQuantity == "3F+N+T") //alterar função para circuitos mono/bi/trifasicos
                    {
                        for (int i = 0; i < 3; i++) //insere os DPS da fase
                        {
                            if (SPD.Circuit == "DPS") InserirBloco("DPS.dwg", new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0));

                            startPt = new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0); //desenha linha que conecta ao barramento principal
                            endPt = new Point3d((lastPoint.X + firstLineLength + secondLineLengthBreaker), (lastPoint.Y + arrowMiddle), 0);
                            busLine.AddBus(startPt, endPt, LineWeight.ByBlock, "Continuous", CircuitsLayer);

                            wiring.AddPhase(new Point3d(((startPt.X + endPt.X) / 2), endPt.Y, 0), LineWeight.ByLayer, WiringLayer, "F"); //desenha fiação

                            content = SPD.ConductorSection;
                            location = new Point3d(((startPt.X + endPt.X) / 2), endPt.Y, 0);
                            descricao.AddDescription(location, white, content, AttachmentPoint.BottomLeft, FontSize, 0, FontName);

                            startPt = new Point3d((lastPoint.X + firstLineLength + (breakerLength) + 43), (lastPoint.Y + arrowMiddle), 0); //desenha linha que conecta ao barramento de terra
                            endPt = new Point3d((lastPoint.X + firstLineLength + (breakerLength)) + 120, (lastPoint.Y + arrowMiddle), 0);
                            busLine.AddBus(startPt, endPt, LineWeight.ByBlock, "Continuous", CircuitsLayer);

                            content = SPD.Description.Remove(13) + "\n" + SPD.Description.Substring(16, 5) + "-275V"; //insere descrição do DPS
                            location = new Point3d((lastPoint.X + firstLineLength + (breakerLength)) + 80, (lastPoint.Y + arrowMiddle), 0);
                            descricao.AddDescription(location, white, content, AttachmentPoint.MiddleCenter, 8, 0, FontName);

                            lastPoint = new Point3d(lastPoint.X, lastPoint.Y + nextCircuit, 0);//Acrescenta o deslocamento para o próximo circuito
                        }

                        //desenha DPS do neutro: 
                        if (SPD.Circuit == "DPS") InserirBloco("DPS.dwg", new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0));

                        neutralSPDLocation = new Point3d((lastPoint.X + firstLineLength + (breakerLength)), (lastPoint.Y + arrowMiddle), 0);

                        neutralBusPos.Add(neutralSPDLocation);

                        startPt = new Point3d((lastPoint.X + firstLineLength + (breakerLength) + 43), (lastPoint.Y + arrowMiddle), 0); //desenha linha que conecta ao barramento de terra
                        endPt = new Point3d((lastPoint.X + firstLineLength + (breakerLength)) + 120, (lastPoint.Y + arrowMiddle), 0);
                        busLine.AddBus(startPt, endPt, LineWeight.ByBlock, "Continuous", CircuitsLayer);

                        content = SPD.Description.Remove(13) + "\n" + SPD.Description.Substring(16, 5) + "-275V"; //insere descrição do DPS 
                        location = new Point3d((lastPoint.X + firstLineLength + (breakerLength)) + 80, (lastPoint.Y + arrowMiddle), 0);
                        descricao.AddDescription(location, white, content, AttachmentPoint.MiddleCenter, 8, 0, FontName);

                        startPt = new Point3d((lastPoint.X + firstLineLength + (breakerLength)) + 120, ((lastPoint.Y - nextCircuit * 3) + arrowMiddle), 0); //desenha o barramento de terra conectado aos dps
                        endPt = new Point3d((lastPoint.X + firstLineLength + (breakerLength)) + 120, ((lastPoint.Y - nextCircuit * 3) + arrowMiddle + nextCircuit * 3), 0);
                        busLine.AddBus(startPt, endPt, LineWeight.LineWeight080, "Continuous", GroundLayer);

                        groundSPDPos = endPt;

                        groundSPDMidPoint = new Point3d(endPt.X, startPt.Y - (startPt.Y - endPt.Y) / 2.0, 0);

                        lastPoint = new Point3d(lastPoint.X, lastPoint.Y + nextCircuit, 0);//Acrescenta o deslocamento para o próximo circuito

                    }


                    descricao.DrawDescription(trans, btr);
                    busLine.DrawBus(trans, btr);
                    wiring.DrawWires(trans, btr);
                }
                catch (System.Exception)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Ocorreu um erro: ");
                    trans.Abort();
                }
            }

            private void DRGroup(ExcelAccess rdExcel, Transaction trans, BlockTableRecord btr, int count, int rowCount, ref Point3d lastPoint) //processa o grupo de DR
            {

                ExcelAccess.Row row;

                string idrCorrente;
                string idrSensibilidade;
                string _idrPolos = "";
                string breakerPolo = "";
                string rdcClass = "Classe A";

                //acresenta espaçamento com circuitos desenhados acima (utilizei expressão condicional ternária)
                lastPoint = (count == 1) ? new Point3d(lastPoint.X, lastPoint.Y - 70, 0) : new Point3d(lastPoint.X, lastPoint.Y - 80, 0);

                row = rdExcel.rows[rowCount - count];

                rdcGroup.AddRDC(row.RCDNumbering, row.RCDPolesQuant, row.RCDCurrent, row.RCDLeakage, rdcClass, "#4", count);

                int pos = int.Parse(row.RCDNumbering) - 1;

                RDCGroups rdcRead = rdcGroup.ReadRDC(pos);


                for (int i = (rowCount - count); i < rowCount; i++)
                {
                    row = rdExcel.rows[i];

                    DrawLinesCircuits(trans, btr, false, ref lastPoint);
                    DrawDescCircuits(rdExcel, trans, btr, i, ref lastPoint);

                    if (i == (rowCount - count)) idrCorrente = row.RCDCurrent;
                    if (i == (rowCount - count)) idrSensibilidade = row.RCDLeakage;
                    if (i == (rowCount - count)) _idrPolos = row.RCDPolesQuant;
                    if (i == (rowCount - count)) breakerPolo = row.CBPolesQuantity;
                }

                DrawBus(rdExcel, trans, btr, false, count, rowCount, ref lastPoint);
                InsertDR(rdExcel, trans, btr, count, rowCount, _idrPolos, ref lastPoint);

                //acresenta espaçamento com circuitos desenhados abaixo
                lastPoint = (count == 1) ? new Point3d(lastPoint.X, lastPoint.Y - 70, 0) : new Point3d(lastPoint.X, lastPoint.Y - 80, 0);


            }

            private void ReadRow(ExcelAccess rdExcel, Transaction trans, BlockTableRecord btr, Point3d lastPt) //verifica os grupos de DR e circuitos sem DR
            {
                int countIDR = 0;
                string RCDNumb = "";
                string description = "";
                string previousRow = "-1";
                ExcelAccess.Row row;
                ExcelAccess.Row row2;

                for (int i = 0; i < (rdExcel.rows.Count); i++)
                {
                    row = rdExcel.rows[i];

                    // contabiliza os grupos de IDRs

                    RCDNumb = row.RCDNumbering;

                    //if ((IDRNumb != previousRow) && (IDRNumb != "0" || IDRNumb != "") && (previousRow != "-1") && (previousRow != "0"))

                    if ((RCDNumb != previousRow) && (!string.IsNullOrEmpty(RCDNumb)) && (previousRow != "-1") && (previousRow != "0"))
                    {
                        //rcdGroups.AddRDC(row.RCDNumbering, row.RCDPolesQuant, row.RCDCurrent, row.RCDLeakage, "Classe A", i);
                        DRGroup(rdExcel, trans, btr, countIDR, i, ref lastPt);
                        countIDR = 0;
                    }

                    int idrNumber;
                    if (previousRow == "0" && int.TryParse(RCDNumb, out idrNumber) && idrNumber > 0)
                    {
                        countIDR = 0;
                    }

                    if (RCDNumb == "0")
                    {
                        DrawLinesCircuits(trans, btr, true, ref lastPt);
                        DrawDescCircuits(rdExcel, trans, btr, i, ref lastPt);
                    }

                    countIDR++;
                    previousRow = row.RCDNumbering;

                }



                for (int j = 0; j < rdExcel.others.Count; j++)
                {
                    row2 = rdExcel.others[j];

                    if (row2.Circuit == "DPS")
                    {
                        InsertSPD(rdExcel, trans, btr, j, ref lastPt);
                    }
                    if (row2.Circuit == "PROTEÇÃO DPS")
                    {

                    }

                }

                DrawBus(rdExcel, trans, btr, true, rdExcel.rows.Count, rdExcel.rows.Count, ref lastPt);

                for (int j = 0; j < rdExcel.others.Count; j++)
                {
                    row2 = rdExcel.others[j];

                    if (row2.Circuit == "ENTRADA DE ENERGIA")
                    {
                        InsertMainBreaker(rdExcel, trans, btr, j, ref lastPt);
                    }
                }

                DrawBox(rdExcel, trans, btr, countIDR, ref lastPt);
                neutralBusPos.Clear(); //limpa a lista para o proximo circuito


            }

            public void Arrow(Transaction trans, BlockTableRecord btr, Point3d startPt, LineWeight lineWeight, string linetype, string layerName)
            {

                Autodesk.AutoCAD.DatabaseServices.Polyline arrow = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                arrow.AddVertexAt(0, new Point2d(startPt.X, (startPt.Y)), 0, 0, 0);
                arrow.AddVertexAt(1, new Point2d(startPt.X, (startPt.Y + arrowMiddle)), 0, 0, 0);
                arrow.AddVertexAt(2, new Point2d((startPt.X + arrowLength), (startPt.Y)), 0, 0, 0);
                arrow.AddVertexAt(3, new Point2d((startPt.X), (startPt.Y - arrowMiddle)), 0, 0, 0);
                arrow.Closed = true;

                arrow.LineWeight = lineWeight;
                arrow.Layer = layerName;
                arrow.Linetype = linetype;

                btr.AppendEntity(arrow);
                trans.AddNewlyCreatedDBObject(arrow, true);
            }

            public static void InserirBloco(string result, Point3d lastPoint)
            {
                // Get the current document and database
                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;


                string relativeFolderPath = "BLOCOS";

                string currentDirectory = Environment.CurrentDirectory;

                string folderPath = Path.Combine(currentDirectory, relativeFolderPath);

                DirectoryInfo directoryInfo = new DirectoryInfo(folderPath);
                FileInfo[] files = directoryInfo.GetFiles("*.dwg");

                foreach (FileInfo file in files)
                {
                    if (result.Equals(file.Name)) result = file.Name;
                }

                string selectedFileName = result;
                string selectedFilePath = Path.Combine(folderPath, selectedFileName);

                // Create a new database and read the DWG into it
                Database acExtDb = new Database(false, true);
                acExtDb.ReadDwgFile(selectedFilePath, System.IO.FileShare.Read, true, "");

                // Create a collection to hold the ObjectId's of the objects we're going to clone
                ObjectIdCollection acObjIdColl = new ObjectIdCollection();

                // Start a transaction in the external database
                using (Transaction acTransExt = acExtDb.TransactionManager.StartTransaction())
                {
                    // Open the Block table for read
                    BlockTable acBlkTblExt;
                    acBlkTblExt = acTransExt.GetObject(acExtDb.BlockTableId,
                                                       OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for read
                    BlockTableRecord acBlkTblRecExt;
                    acBlkTblRecExt = acTransExt.GetObject(acBlkTblExt[BlockTableRecord.ModelSpace],
                                                          OpenMode.ForRead) as BlockTableRecord;

                    // Add all the objects from the Model space of the external database to the ObjectIdCollection
                    foreach (ObjectId id in acBlkTblRecExt)
                    {
                        acObjIdColl.Add(id);
                    }

                    // Save the changes made to the external database
                    acTransExt.Commit();
                }

                // Lock the current document
                using (DocumentLock acLckDocCur = acDoc.LockDocument())
                {
                    // Start a transaction
                    using (Transaction acTransCur = acCurDb.TransactionManager.StartTransaction())
                    {
                        // Open the Block table for read
                        BlockTable acBlkTblCur;
                        acBlkTblCur = acTransCur.GetObject(acCurDb.BlockTableId,
                                                           OpenMode.ForRead) as BlockTable;

                        // Open the Block table record Model space for write
                        BlockTableRecord acBlkTblRecCur;
                        acBlkTblRecCur = acTransCur.GetObject(acBlkTblCur[BlockTableRecord.ModelSpace],
                                                              OpenMode.ForWrite) as BlockTableRecord;

                        IdMapping acIdMap = new IdMapping();
                        acExtDb.WblockCloneObjects(acObjIdColl, acBlkTblRecCur.ObjectId, acIdMap,
                                                   DuplicateRecordCloning.Ignore, false);

                        //lastPoint.Y -= deslocamento;

                        // Calculate the vector for the move operation
                        Vector3d moveVec = new Point3d(lastPoint.X, lastPoint.Y, 0) - Point3d.Origin;



                        // Move the cloned objects to the new location
                        foreach (IdPair idPair in acIdMap)
                        {
                            // Open the cloned object for write
                            Entity ent = acTransCur.GetObject(idPair.Value, OpenMode.ForWrite) as Entity;
                            if (ent != null)
                            {
                                // Move the entity
                                ent.TransformBy(Matrix3d.Displacement(moveVec));
                            }
                        }

                        // Save the changes made to the current database
                        acTransCur.Commit();
                    }

                    // Unlock the document
                }

                // Dispose of the external database
                acExtDb.Dispose();
            }

        
        public void Terminate()
        {
            throw new NotImplementedException();
        }

        
    }
}
