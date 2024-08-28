using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToAutoCAD.Entities
{
    public class TableLoads
    {
        public void CreateTable(DataGridView dgv, string boardName) { 
            CadManager cadManager = new CadManager();
            Document doc = cadManager.GetDocument();
            Database db = doc.Database;
            Editor edt = doc.Editor;

            PromptPointResult pr = edt.GetPoint("\nInsira o ponto de inserção: ");

            if (pr.Status == PromptStatus.OK)

            {
                Table tb = new Table();
                tb.TableStyle = db.Tablestyle;
                tb.NumRows = dgv.Rows.Count;
                tb.NumColumns = dgv.Columns.Count;
                tb.SetRowHeight(30);
                tb.SetColumnWidth(200);
                tb.Position = pr.Value;

                

                              
                // Insert a row above the headers
                tb.InsertRows(0, 1, 1);

                // Set the text for the new row
                tb.SetTextString(0, 0, boardName);

                // Merge the cells of the new row to span across all columns
                CellRange headerTop = CellRange.Create(tb, 0, 0, 0, tb.NumColumns - 1);
                tb.MergeCells(headerTop);

                // Set the alignment for the new row
                tb.SetAlignment(0, 0, CellAlignment.MiddleCenter);

                // Set the height for the new row
                tb.SetRowHeight(0, 60);
                               
                                

                // Populate the header and adjust column widths as before
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    tb.SetTextHeight(13, i);
                    tb.Cells[1, i].TextString = dgv.Columns[i].HeaderText;

                    double maxLength = dgv.Columns[i].HeaderText.Length;

                    for (int j = 0; j < dgv.Rows.Count; j++)
                    {
                        DataGridViewCell cell = dgv.Rows[j].Cells[i];
                        if (cell.Value != null)
                        {
                            double textLenght = cell.Value.ToString().Length;

                            if (textLenght > maxLength)
                            {
                                maxLength = textLenght;
                            }
                        }
                    }
                    tb.SetColumnWidth(i, maxLength * 11);
                }

                // Populate the cells as before
                for (int i = 0; i < dgv.Rows.Count - 1; i++)
                {
                    DataGridViewRow row = dgv.Rows[i];
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        tb.SetTextHeight(i + 2, j, 12);
                        if (row.Cells[j].Value == null)
                            tb.SetTextString(i + 2, j, " ");
                        else
                            tb.SetTextString(i + 2, j, row.Cells[j].Value.ToString());
                        tb.SetAlignment(i + 2, j, CellAlignment.MiddleCenter);
                    }
                }

                // Ensure the new layout is generated
                tb.GenerateLayout();

                var rowHeader = tb.Rows[1];
                if (rowHeader.IsMerged.HasValue && rowHeader.IsMerged.Value)
                {
                    tb.UnmergeCells(rowHeader);
                }

                Transaction tr = doc.TransactionManager.StartTransaction();
                using (tr)
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();
                    BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    btr.AppendEntity(tb);
                    tr.AddNewlyCreatedDBObject(tb, true);

                    tr.Commit();
                }


            }



        }


    }
}
