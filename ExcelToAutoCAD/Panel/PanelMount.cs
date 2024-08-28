using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.IO;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelToAutoCAD.ExcelAccess;

namespace ExcelToAutoCAD.Panel
{
    public class PanelMount
    {

        Dictionary<string, List<string>> attributesToUpdate = new Dictionary<string, List<string>>();

        ExcelAccess ReadExcel = new ExcelAccess();

        public string pathFile { get; set; }
        public string sheetName { get; set; }

        List<string> blockType = new List<string>();

        public PanelMount(string pathFile, string sheetName)
        {
            this.pathFile = pathFile;
            this.sheetName = sheetName;
        }

        public void GeneratePanel()
        {
            CadManager cadManager = new CadManager();
            Document doc = cadManager.GetDocument();
            Editor ed = doc.Editor;

            PromptPointResult pr = ed.GetPoint("\n\nInsira o ponto na tela: ");

            if (pr.Status == PromptStatus.OK)
            {
                Point3d insertionPoint = pr.Value;
                double desloc = 0;

                ReadBlockExcel(pathFile, sheetName);

                string relativeFolderPath = "BLOCOS/Esquema de Montagem";

                string currentDirectory = Environment.CurrentDirectory;
                                

                for (int i = 0; i < (ReadExcel.rows.Count - 1) ; i++)
                {
                    string blockName = "";

                    if (blockType[i] == "UNIPOLAR")
                    {
                        blockName = "Dis_Mono_1";
                        relativeFolderPath = "BLOCOS/Esquema de Montagem/MONO.dwg";
                        desloc = 40;
                    }
                    else if (blockType[i] == "BIPOLAR")
                    {
                        blockName = "Dis_bi_";
                        relativeFolderPath = "BLOCOS/Esquema de Montagem/BIFASICO.dwg";
                        desloc = 80;
                    }
                    else if (blockType[i] == "TRIPOLAR")
                    {
                        blockName = "Disjuntor_Trifasico";
                        relativeFolderPath = "BLOCOS/Esquema de Montagem/TRIFASICO.dwg";
                        desloc = 120;
                    }
                    else if (blockType[i] == "")
                    {
                        
                    }

                    string folderPath = Path.Combine(currentDirectory, relativeFolderPath);

                    InsertBlock(folderPath, blockName, insertionPoint);
                    insertionPoint = new Point3d(insertionPoint.X + desloc, insertionPoint.Y, insertionPoint.Z);
                }
            }



        }

        private void InsertBlock(string blockPath, string blockName, Point3d insertionPoint)
        {
            CadManager cadManager = new CadManager();

            var doc = cadManager.GetDocument();
            var db = doc.Database;
            var ed = doc.Editor;
            using (var tr = db.TransactionManager.StartTransaction())
            {

                Application.DocumentManager.MdiActiveDocument.LockDocument();
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                ObjectId btrId = bt.Has(blockName) ?
                    bt[blockName] :
                    ImportBlock(db, blockName, blockPath);
                if (btrId.IsNull)
                {
                    ed.WriteMessage($"\nBlock '{blockName}' not found.");
                    return;
                }

                var cSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                var br = new BlockReference(insertionPoint, btrId);
                cSpace.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);

                // Adicione os atributos ao bloco
                AddAttributesToBlock(tr, br, btrId);

                tr.Commit();

                DynamicBlockProps(br);
            }
        }

        private void AddAttributesToBlock(Transaction tr, BlockReference br, ObjectId btrId)
        {
            var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            if (btr.HasAttributeDefinitions)
            {
                foreach (ObjectId id in btr)
                {
                    if (id.ObjectClass.DxfName == "ATTDEF")
                    {
                        var attDef = (AttributeDefinition)tr.GetObject(id, OpenMode.ForRead);
                        var attRef = new AttributeReference();
                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                        attRef.TextString = attDef.TextString;
                        br.AttributeCollection.AppendAttribute(attRef);
                        tr.AddNewlyCreatedDBObject(attRef, true);
                    }
                }
            }
        }

        private ObjectId ImportBlock(Database destDb, string blockName, string sourceFileName)
        {
            if (System.IO.File.Exists(sourceFileName))
            {
                using (var sourceDb = new Database(false, true))
                {
                    try
                    {
                        Application.DocumentManager.MdiActiveDocument.LockDocument();
                        sourceDb.ReadDwgFile(sourceFileName, FileOpenMode.OpenForReadAndAllShare, true, "");

                        var id = ObjectId.Null;
                        using (var tr = new OpenCloseTransaction())
                        {
                            var bt = (BlockTable)tr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false);

                            if (bt.Has(blockName))
                                id = bt[blockName];
                        }
                        if (!id.IsNull)
                        {
                            var blockIds = new ObjectIdCollection();
                            blockIds.Add(id);
                            var mapping = new IdMapping();
                            sourceDb.WblockCloneObjects(blockIds, destDb.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
                            if (mapping[id].IsCloned)
                                return mapping[id].Value;
                        }
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        CadManager cadManager = new CadManager();
                        var doc = cadManager.GetDocument();
                        var ed = doc.Editor;

                        ed.WriteMessage("\nError during copy: " + ex.Message + "\n" + ex.StackTrace);
                    }
                }
            }
            return ObjectId.Null;
        }



        public void DynamicBlockProps(BlockReference br)
        {
            CadManager cadManager = new CadManager();

            Document doc = cadManager.GetDocument();
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Transaction tr = db.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);



                UpdateDynamicBlockAttributes(db, tr, br, attributesToUpdate);

                tr.Commit();
            }
        }

        private void UpdateDynamicBlockAttributes(Database db, Transaction tr, BlockReference br, Dictionary<string, List<string>> attributeValues)
        {
            if (br != null)
            {
                AttributeCollection attColl = br.AttributeCollection;

                foreach (ObjectId attId in attColl)
                {
                    AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);

                    // Verifica se a tag do atributo existe no dicionário de valores
                    if (attributeValues.ContainsKey(attRef.Tag))
                    {
                        // Obtém a lista de valores para essa tag
                        List<string> values = attributeValues[attRef.Tag];

                        // Se a lista de valores não estiver vazia, atualize o atributo com o primeiro valor da lista
                        if (values.Count > 0)
                        {
                            attRef.UpgradeOpen();
                            attRef.TextString = values[0];
                            attRef.DowngradeOpen();

                            // Remove o valor atual da lista para evitar que seja usado novamente
                            values.RemoveAt(0);
                        }
                    }
                }
            }
        }

        public void ReadBlockExcel(string pathFile, string sheetName)
        {
            ReadExcel.ReadColumn(pathFile, sheetName);

            ExcelAccess.Row row = new ExcelAccess.Row();

            for (int i = 0; i < (ReadExcel.rows.Count-2); i++)
            {
                row = ReadExcel.rows[i];

                blockType.Add(row.CBPolesQuantity);

                if (!attributesToUpdate.ContainsKey("00"))
                {
                    attributesToUpdate["00"] = new List<string>();
                }
                if (!attributesToUpdate.ContainsKey("00A"))
                {
                    attributesToUpdate["00A"] = new List<string>();
                }
                if (!attributesToUpdate.ContainsKey("#0,0"))
                {
                    attributesToUpdate["#0,0"] = new List<string>();
                }
                if (!attributesToUpdate.ContainsKey("DESCRIÇÃO"))
                {
                    attributesToUpdate["DESCRIÇÃO"] = new List<string>();
                }

                attributesToUpdate["00"].Add(row.Circuit.ToString());
                attributesToUpdate["00A"].Add(row.CBCurrent.ToString());
                attributesToUpdate["#0,0"].Add(row.ConductorSection.ToString());

                string LetterLower;
                string firstLetterUpper;
                string newLine = "";

                if (!string.IsNullOrEmpty(row.Description.ToString()))
                {
                    LetterLower = row.Description.ToString().ToLower();
                    firstLetterUpper = char.ToUpper(LetterLower[0]) + LetterLower.Substring(1);

                    if (firstLetterUpper.Length > 13)
                    {

                        for (int j = 0; j < firstLetterUpper.Length; j++)
                        {
                            newLine += firstLetterUpper[j];
                            if (j > 0 && (j + 1) % 13 == 0 && j != firstLetterUpper.Length - 1)
                            {
                                if (firstLetterUpper[j + 1] != ' ')
                                {
                                    newLine += "-\n";
                                }
                                else
                                {
                                    newLine += "\n";
                                }
                            }
                        }
                    }
                    else
                    {
                        newLine = firstLetterUpper;
                    }

                }
                else
                {
                    newLine = "";
                }

                attributesToUpdate["DESCRIÇÃO"].Add(newLine);

            }

            for (int j = 0; j < ReadExcel.others.Count; j++)
            {
                row = ReadExcel.others[j];
                if (row.Circuit == "ENTRADA DE ENERGIA")
                {
                    blockType.Add(row.CBPolesQuantity);

                    if (!attributesToUpdate.ContainsKey("00"))
                    {
                        attributesToUpdate["00"] = new List<string>();
                    }
                    if (!attributesToUpdate.ContainsKey("00A"))
                    {
                        attributesToUpdate["00A"] = new List<string>();
                    }
                    if (!attributesToUpdate.ContainsKey("#0,0"))
                    {
                        attributesToUpdate["#0,0"] = new List<string>();
                    }
                    if (!attributesToUpdate.ContainsKey("DESCRIÇÃO"))
                    {
                        attributesToUpdate["DESCRIÇÃO"] = new List<string>();
                    }

                    attributesToUpdate["00"].Add(row.Circuit.ToString());
                    attributesToUpdate["00A"].Add(row.CBCurrent.ToString());
                    attributesToUpdate["#0,0"].Add(row.ConductorSection.ToString());
                    attributesToUpdate["DESCRIÇÃO"].Add("Geral");
                }
            }

        }


    }
}


