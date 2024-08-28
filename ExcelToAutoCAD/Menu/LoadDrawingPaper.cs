using Autodesk.AutoCAD.ApplicationServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD.Menu
{
    
    public class LoadDrawingPaper
    {
       public string filePath { get; set; } = "";

        public void OpenDrawing()
        {
            DocumentCollection docMgr = Application.DocumentManager;

            if(File.Exists(filePath))
            {
                docMgr.Open(filePath, false);
            }
            else
            {
                docMgr.MdiActiveDocument.Editor.WriteMessage("Arquivo " + filePath + "não encontrado");
            }

        }

    }
}
