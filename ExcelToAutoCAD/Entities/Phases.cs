using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD.Entities
{
    internal class Phases
    {
        const double phaseLength = 15;
        public Point3d StartPoint { get; set; }
        public Point3d EndPoint { get; set; }

        public int ColoIndex { get; set; }

        public LineWeight LineWeight_ { get; set; }

        public List<Phases> phaseLines { get; set; } = new List<Phases>();

        public Phases() { }

        public Phases(Point3d startPoint, Point3d endPoint, int colorIndex, LineWeight lineWeight)
        {
            StartPoint = startPoint;
            EndPoint = endPoint;
            ColoIndex = colorIndex;
            LineWeight_ = lineWeight;
        }
        public void AddPhase(Point3d startPoint, Point3d endPoint, int colorIndex, LineWeight lineWeight)
        {
            StartPoint = startPoint;
            EndPoint = endPoint;
            ColoIndex = colorIndex;
            LineWeight_ = lineWeight;            
        }


        public void DrawBipolar(Transaction trans, BlockTableRecord btr)
        {

            Line neutralLine = new Line();
            neutralLine.StartPoint = new Point3d(StartPoint.X, StartPoint.Y + phaseLength / 2, 0);
            neutralLine.EndPoint = new Point3d(EndPoint.X-5, StartPoint.Y + phaseLength / 2, 0);
            neutralLine.ColorIndex = ColoIndex;
            neutralLine.LineWeight = LineWeight_;
            btr.AppendEntity(neutralLine);
            trans.AddNewlyCreatedDBObject(neutralLine, true);

            Line neutralLine2 = new Line();
            neutralLine2.StartPoint = new Point3d(StartPoint.X, StartPoint.Y + phaseLength / 2, 0);
            neutralLine2.EndPoint = new Point3d(EndPoint.X, StartPoint.Y - phaseLength / 2, 0);
            neutralLine2.ColorIndex = ColoIndex;
            neutralLine2.LineWeight = LineWeight_;
            btr.AppendEntity(neutralLine2);
            trans.AddNewlyCreatedDBObject(neutralLine2, true);

            Line phaseLine = new Line();
            phaseLine.StartPoint = new Point3d(StartPoint.X + 5, StartPoint.Y + phaseLength / 2, 0);
            phaseLine.EndPoint = new Point3d(EndPoint.X + 5, StartPoint.Y - phaseLength / 2, 0);
            phaseLine.ColorIndex = ColoIndex;
            phaseLine.LineWeight = LineWeight_;
            btr.AppendEntity(phaseLine);
            trans.AddNewlyCreatedDBObject(phaseLine, true);

        }

        public void DrawTetrapolar(Transaction trans, BlockTableRecord btr)
        {
            StartPoint = new Point3d(StartPoint.X, StartPoint.Y + phaseLength / 2, 0);
            EndPoint = new Point3d(EndPoint.X, EndPoint.Y - phaseLength / 2, 0);

            Line neutralLine = new Line(StartPoint, new Point3d(StartPoint.X - 5, StartPoint.Y, 0));
            neutralLine.ColorIndex = ColoIndex;
            neutralLine.LineWeight = LineWeight_;
            btr.AppendEntity(neutralLine);
            trans.AddNewlyCreatedDBObject(neutralLine, true);

            for (int i = 0; i < 4; i++)
            {
                phaseLines.Add(new Phases(StartPoint, EndPoint, ColoIndex, LineWeight_));
                StartPoint = new Point3d(StartPoint.X + 5, StartPoint.Y, StartPoint.Z);
                EndPoint = new Point3d(EndPoint.X + 5, EndPoint.Y, EndPoint.Z);
            }


            foreach (Phases p in phaseLines)
            {
                Line phaseLine = new Line(p.StartPoint, p.EndPoint);
                phaseLine.ColorIndex = p.ColoIndex;
                phaseLine.LineWeight = p.LineWeight_;
                btr.AppendEntity(phaseLine);
                trans.AddNewlyCreatedDBObject(phaseLine, true);

            }
        }

        public void DrawPhase(Transaction trans, BlockTableRecord btr)
        {
            StartPoint = new Point3d(StartPoint.X, StartPoint.Y + phaseLength / 2, 0);
            EndPoint = new Point3d(EndPoint.X, EndPoint.Y - phaseLength / 2, 0);

            Line Phase = new Line();
            Phase.StartPoint = new Point3d(StartPoint.X + 5, StartPoint.Y, StartPoint.Z);
            Phase.EndPoint = new Point3d(EndPoint.X + 5, EndPoint.Y, EndPoint.Z);

        }





    }
}
