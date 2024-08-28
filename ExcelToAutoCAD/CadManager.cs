using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD
{
    public class CadManager
    {
        // Campos para armazenar instâncias do Document, Database e Editor
        private Document doc;
        private Database db;
        private Editor edt;
        

        // Construtor da classe CadManager
        public CadManager()
        {
            // Inicialize as instâncias do Document, Database e Editor
            doc = Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            edt = doc.Editor;
            
        }

        // Método para acessar o Document
        public Document GetDocument()
        {
            return doc;
        }

        // Método para acessar o Database
        public Database GetDatabase()
        {
            return db;
        }

        // Método para acessar o Editor
        public Editor GetEditor()
        {
            return edt;
        }
        

    }
}
