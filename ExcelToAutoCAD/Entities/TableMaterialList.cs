using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using static ExcelToAutoCAD.ExcelAccess;

namespace ExcelToAutoCAD.Entities
{
    public class TableMaterialList
    {
        public void CreateTable(DataGridView dgv)
        {
            CadManager cadManager = new CadManager();
            Document doc = cadManager.GetDocument();
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptPointResult pr = ed.GetPoint("\nInsira o ponto de inserção: ");

            if (pr.Status == PromptStatus.OK)

            {
                Table tb = new Table();
                tb.TableStyle = db.Tablestyle;
                tb.NumRows = dgv.Rows.Count;
                tb.NumColumns = dgv.Columns.Count;
                tb.SetRowHeight(30);
                tb.SetColumnWidth(200);
                tb.Position = pr.Value;
                
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    tb.SetTextHeight(12, i);
                   // tb.SetTextString(0, i, dgv.Columns[i].HeaderText);
                    tb.Cells[0, i].TextString = dgv.Columns[i].HeaderText;

                    double maxLength = dgv.Columns[i].HeaderText.Length;

                    for(int j = 0; j< dgv.Rows.Count; j++)
                    {
                        DataGridViewCell cell = dgv.Rows[j].Cells[i];
                        if(cell.Value != null)
                        {
                            double textLenght = cell.Value.ToString().Length;

                            if(textLenght > maxLength)
                            {
                                maxLength = textLenght;
                            }
                        }
                    }
                    tb.SetColumnWidth(i, maxLength * 11);
                }
                


                // Use a nested loop to add and format each cell
                for (int i = 0; i < dgv.Rows.Count -1; i++)
                {
                    

                    DataGridViewRow row = dgv.Rows[i];
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        tb.SetTextHeight(i+1, j, 12);
                        if (row.Cells[j].Value == null)
                            tb.SetTextString(i + 1, j, " ");
                        else
                            tb.SetTextString(i + 1, j, row.Cells[j].Value.ToString());
                        tb.SetAlignment(i+1, j, CellAlignment.MiddleCenter);
                    }
                }               
                tb.GenerateLayout();

                var rowHeader = tb.Rows[0];
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




