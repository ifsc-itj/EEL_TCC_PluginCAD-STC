using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD.PrintMgr
{
    public class PrintDraw
    {
        public List<string> plotterDevices { get;} = new List<string>();
        public List<string> paperSize { get;} = new List<string>();

        public string plotDeviceName { get; set; } = "DWG To PDF.pc3";       
        public string paperName { get; set; } = "ISO_full_bleed_A1_(594.00_x_841.00_MM)";


        public void LayoutCreate()
        {
            CadManager cadManager = new CadManager();
            Document doc = cadManager.GetDocument();
            Database db = doc.Database;
            Editor edt = doc.Editor;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                Application.DocumentManager.MdiActiveDocument.LockDocument();

                LayoutManager layoutMgr = LayoutManager.Current;

                // Criar o novo layout com as configurações padrão
                ObjectId objID = layoutMgr.CreateLayout("STC - Layout");

                // Abrir o layout
                Layout acLayout = trans.GetObject(objID, OpenMode.ForRead) as Layout;

                acLayout.CopyFrom(db);
                
                // Definir o layout como o atual se ainda não estiver selecionado
                if (!acLayout.TabSelected)
                {
                    layoutMgr.CurrentLayout = acLayout.LayoutName;
                }

                trans.Commit();
            }
        }

        public void PlotterList()
        {
            CadManager cadManager = new CadManager();
            Document doc = cadManager.GetDocument();
                                 
            foreach (string plotDevice in PlotSettingsValidator.Current.GetPlotDeviceList())
            {
                // Output the names of the available plotter devices
                plotterDevices.Add(plotDevice);
            }
        }

        public void PaperSize()
        {
            CadManager cadManager = new CadManager();
            Document doc = cadManager.GetDocument();

            using (PlotSettings plSet = new PlotSettings(true))
            {
                PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                // Define a configuração de plotagem (Plotter e tamanho da página)
                acPlSetVdr.SetPlotConfigurationName(plSet, plotDeviceName, paperName);
                                
                foreach (string mediaName in acPlSetVdr.GetCanonicalMediaNameList(plSet))
                {                                        
                    paperSize.Add(mediaName);
                }
            }
        }

    }
}
