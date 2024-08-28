using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.IO;

namespace ExcelToAutoCAD
{
    public class BlockReference
    {
        [CommandMethod("insertBlock")]
        public void InsertBlock()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            string folderPath = "C:\\Users\\gusta\\OneDrive - Instituto Federal de Santa Catarina\\9º Fase\\Projeto Integrador 3\\Plugin AutoCad\\Projeto\\CAD\\Blocos"; // Especifique o caminho da pasta desejada

            DirectoryInfo directoryInfo = new DirectoryInfo(folderPath);
            FileInfo[] files = directoryInfo.GetFiles("*.dwg");

            PromptKeywordOptions keywordOptions = new PromptKeywordOptions("Selecione um arquivo de bloco");
            foreach (FileInfo file in files)
            {
                keywordOptions.Keywords.Add(file.Name);
            }

            PromptResult result = ed.GetKeywords(keywordOptions);
            if (result.Status != PromptStatus.OK) return;

            string selectedFileName = result.StringResult;
            string selectedFilePath = Path.Combine(folderPath, selectedFileName);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                if (!File.Exists(selectedFilePath)) return;

                using (Database blockDb = new Database(false, true))
                {
                    blockDb.ReadDwgFile(selectedFilePath, FileShare.Read, true, "");

                    ObjectIdCollection ids = new ObjectIdCollection();
                    using (Transaction blockTr = blockDb.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord blockModelSpace = (BlockTableRecord)blockTr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(blockDb), OpenMode.ForRead);

                        foreach (ObjectId objectId in blockModelSpace)
                        {
                            Entity entity = (Entity)blockTr.GetObject(objectId, OpenMode.ForRead);
                            if (entity != null)
                            {
                                ids.Add(objectId);
                            }
                        }

                        blockTr.Commit();
                    }

                    IdMapping mapping = new IdMapping();
                    db.WblockCloneObjects(ids, modelSpace.Id, mapping, DuplicateRecordCloning.Ignore, false);

                    foreach (IdPair pair in mapping)
                    {
                        if (!pair.IsPrimary) continue;

                        Autodesk.AutoCAD.DatabaseServices.BlockReference blockRef = new Autodesk.AutoCAD.DatabaseServices.BlockReference(Point3d.Origin, pair.Value);
                        modelSpace.AppendEntity(blockRef);
                        tr.AddNewlyCreatedDBObject(blockRef, true);
                    }
                }

                tr.Commit();
            }
        }

        [CommandMethod("insertBlock2")]
        public void InsertBlock2()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            string folderPath = "C:\\Users\\gusta\\OneDrive - Instituto Federal de Santa Catarina\\9º Fase\\Projeto Integrador 3\\Plugin AutoCad\\Projeto\\CAD\\Blocos";
            DirectoryInfo directoryInfo = new DirectoryInfo(folderPath);
            FileInfo[] files = directoryInfo.GetFiles("*.dwg");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                foreach (FileInfo file in files)
                {
                    using (Database sourceDb = new Database(false, true))
                    {
                        sourceDb.ReadDwgFile(file.FullName, FileShare.Read, true, "");

                        using (Transaction sourceTr = sourceDb.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord sourceModelSpace = (BlockTableRecord)sourceTr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(sourceDb), OpenMode.ForRead);

                            foreach (ObjectId sourceId in sourceModelSpace)
                            {
                                Entity sourceEntity = sourceTr.GetObject(sourceId, OpenMode.ForRead) as Entity;
                                if (sourceEntity != null)
                                {
                                    Entity clonedEntity = (Entity)sourceEntity.Clone();
                                    modelSpace.AppendEntity(clonedEntity);
                                    tr.AddNewlyCreatedDBObject(clonedEntity, true);
                                }
                            }

                            sourceTr.Commit();
                        }
                    }
                }

                tr.Commit();
            }
        }


        [CommandMethod("CopyObjectsBetweenDatabases", CommandFlags.Session)]
        public static void CopyObjectsBetweenDatabases()
        {
        ObjectIdCollection acObjIdColl = new ObjectIdCollection();
        // Get the current document and database
        Document acDoc = Application.DocumentManager.MdiActiveDocument;
        Database acCurDb = acDoc.Database;
        // Lock the current document
        using (DocumentLock acLckDocCur = acDoc.LockDocument())
        {
            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table record for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;
                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;
                // Create a circle that is at (0,0,0) with a radius of 5
                Circle acCirc1 = new Circle();
                acCirc1.SetDatabaseDefaults();
                acCirc1.Center = new Point3d(0, 0, 0);
                acCirc1.Radius = 5;
                // Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acCirc1);
                acTrans.AddNewlyCreatedDBObject(acCirc1, true);
                // Create a circle that is at (0,0,0) with a radius of 7
                Circle acCirc2 = new Circle();
                acCirc2.SetDatabaseDefaults();
                acCirc2.Center = new Point3d(0, 0, 0);
                acCirc2.Radius = 7;
                // Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acCirc2);
                acTrans.AddNewlyCreatedDBObject(acCirc2, true);
                // Add all the objects to copy to the new document
                acObjIdColl = new ObjectIdCollection();
                acObjIdColl.Add(acCirc1.ObjectId);
                acObjIdColl.Add(acCirc2.ObjectId);
                // Save the new objects to the database
                acTrans.Commit();
            }
            // Unlock the document
        }
        // Change the file and path to match a drawing template on your workstation
        string sLocalRoot = Application.GetSystemVariable("LOCALROOTPREFIX") as string;
        string sTemplatePath = sLocalRoot + "Template\\acad.dwt";
        // Create a new drawing to copy the objects to
        DocumentCollection acDocMgr = Application.DocumentManager;
        Document acNewDoc = acDocMgr.Add(sTemplatePath);
        Database acDbNewDoc = acNewDoc.Database;
        // Lock the new document
        using (DocumentLock acLckDoc = acNewDoc.LockDocument())
        {
            // Start a transaction in the new database
            using (Transaction acTrans = acDbNewDoc.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTblNewDoc;
                acBlkTblNewDoc = acTrans.GetObject(acDbNewDoc.BlockTableId,
                                                   OpenMode.ForRead) as BlockTable;
                // Open the Block table record Model space for read
                BlockTableRecord acBlkTblRecNewDoc;
                acBlkTblRecNewDoc = acTrans.GetObject(acBlkTblNewDoc[BlockTableRecord.ModelSpace],
                                                   OpenMode.ForRead) as BlockTableRecord;
                // Clone the objects to the new database
                IdMapping acIdMap = new IdMapping();
                acCurDb.WblockCloneObjects(acObjIdColl, acBlkTblRecNewDoc.ObjectId, acIdMap,
                                           DuplicateRecordCloning.Ignore, false);


                    // Iterate through the IdMapping to get the new Ids in the new database
                    foreach (IdPair idPair in acIdMap)
                    {
                        // Open the cloned object for write
                        Entity ent = acTrans.GetObject(idPair.Value, OpenMode.ForWrite) as Entity;
                        if (ent != null)
                        {
                            // Check if the entity is a Circle
                            if (ent is Circle)
                            {
                                Circle circle = ent as Circle;

                                // Calculate the vector for the move operation
                                Vector3d moveVec = new Point3d(50, 50, 0) - circle.Center;

                                // Move the circle
                                circle.TransformBy(Matrix3d.Displacement(moveVec));
                            }
                            // Include other entity types if needed
                        }
                    }
                    // Save the copied objects to the database
                    acTrans.Commit();
            }
            // Unlock the document
        }
        // Set the new document current
        acDocMgr.MdiActiveDocument = acNewDoc;
         }


        [CommandMethod("ImportObjectsFromDWG", CommandFlags.Session)]
        public static void ImportObjectsFromDWG()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor edt = acDoc.Editor;

            PromptOpenFileOptions openFileOpts = new PromptOpenFileOptions("Selecione o arquivo: ");
            openFileOpts.Filter = "DWG files (*.dwg)|*.dwg";
            PromptFileNameResult openFileRes = edt.GetFileNameForOpen(openFileOpts);
            if (openFileRes.Status != PromptStatus.OK)
                return;
            string sExternalDWG = openFileRes.StringResult;

            //string sExternalDWG = "C:\\Users\\gusta\\OneDrive - Instituto Federal de Santa Catarina\\9º Fase\\Projeto Integrador 3\\Plugin AutoCad\\Projeto\\CAD\\Blocos\\1xDIN.dwg";

            // Create a new database and read the DWG into it
            Database acExtDb = new Database(false, true);
            acExtDb.ReadDwgFile(sExternalDWG, System.IO.FileShare.Read, true, "");

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

                    // Calculate the vector for the move operation
                    Vector3d moveVec = new Point3d(50, 50, 0) - Point3d.Origin;

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

    }
}
