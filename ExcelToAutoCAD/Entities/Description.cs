using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ExcelToAutoCAD.Entities
{
    internal class Description
    {
        public Point3d Location { get; set; }
        public int ColorIndex { get; set; }

        public string mTextContent { get; set; }

        public AttachmentPoint Attachment { get; set; }
        public double TextHeight { get; set; }

        public string FontName { get; set; }

        public double Rotation { get; set; }
        public List<Description> DescricaoList { get; set; } = new List<Description>();

        public Description()
        {
        }
        public Description(Point3d location, int colorIndex, string mTextContent, AttachmentPoint attachment, double textHeight, double rotation, string FontName)
        {
            Location = location;
            ColorIndex = colorIndex;
            this.mTextContent = mTextContent;
            Attachment = attachment;
            TextHeight = textHeight;
            this.FontName = FontName;
            Rotation = rotation;
        }

        public void AddDescription(Point3d location, int colorIndex, string mTextContent, AttachmentPoint attachment, double textHeight, double rotation, string FontName)
        {
            DescricaoList.Add(new Description(location, colorIndex, mTextContent, attachment, textHeight, rotation, FontName));
        }


        public void DrawDescription(Transaction trans, BlockTableRecord btr)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (trans = db.TransactionManager.StartTransaction())
            {
                try
                {

                    foreach (Description descricao in DescricaoList)
                    {                                                                 

                        MText mText = new MText();

                        mText.Attachment = descricao.Attachment;
                        mText.Contents =  $"\\f{descricao.FontName}; { descricao.mTextContent}";
                        mText.TextHeight = descricao.TextHeight;
                        mText.ColorIndex = descricao.ColorIndex;                        
                        mText.Location = new Point3d(descricao.Location.X, descricao.Location.Y, descricao.Location.Z);
                        mText.Rotation = descricao.Rotation;

                        btr.AppendEntity(mText);
                        trans.AddNewlyCreatedDBObject(mText, true);
                        
                    }                                       
                        trans.Commit();
                }
                catch (System.Exception ex)
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Ocorreu um erro: ");
                    trans.Abort();
                }

            }
        }

    }
}
