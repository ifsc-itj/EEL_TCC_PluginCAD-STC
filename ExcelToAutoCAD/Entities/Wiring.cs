using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;



namespace ExcelToAutoCAD.Entities
{
    public class Wiring
    {
        const double phaseLength = 15;
        public Point3d StartPoint { get; set; }
        
        public LineWeight LineWeight { get; set; }

        public string LayerName { get; set; }

        public string WireType { get; set; }

        private double Spacing = 5.0;

        
        public List<Wiring> wiringP { get; set; } = new List<Wiring>();

        public Wiring() { }

        public Wiring(Point3d startPoint, LineWeight lineWeight, string layerName, string wireType)
        {
            StartPoint = startPoint;
            LineWeight = lineWeight;
            LayerName = layerName;
            WireType = wireType;
            
        }

        public void AddPhase(Point3d startPoint, LineWeight lineWeight, string layerName, string wireType)
        {
            wiringP.Add(new Wiring(startPoint, lineWeight, layerName, wireType));
        }

        public void DrawWires(Transaction trans, BlockTableRecord btr)
        {
            foreach (Wiring w in wiringP)
            {
                if(w.WireType == "F")
                {
                    LiveWiring(trans, btr, w, 0);
                }
                if (w.WireType == "2F")
                {
                    LiveWiring(trans, btr, w, -2.5);
                    LiveWiring(trans, btr, w, 2.5);
                }
                if (w.WireType == "3F")
                {
                    LiveWiring(trans, btr, w, -5);
                    LiveWiring(trans, btr, w, 0);
                    LiveWiring(trans, btr, w, 5);
                }
                if (w.WireType == "N")
                {
                    NeutralWiring(trans, btr, w, 0);
                }
                if(w.WireType == "T")
                {
                    GroundWiring(trans, btr, w, 0);
                }
                if(w.WireType == "F+N+T")
                {
                    NeutralWiring(trans, btr, w, 0);
                    LiveWiring(trans, btr, w, 5);
                    GroundWiring(trans, btr, w, 15);
                }
                if(w.WireType == "2F+N+T")
                {
                    NeutralWiring(trans, btr, w, 0);
                    LiveWiring(trans, btr, w, 5);
                    LiveWiring(trans, btr, w, 10);
                    GroundWiring(trans, btr, w, 20);
                }
                if (w.WireType == "3F+N+T")
                {
                    NeutralWiring(trans, btr, w, 0);
                    LiveWiring(trans, btr, w, 5);
                    LiveWiring(trans, btr, w, 10);
                    LiveWiring(trans, btr, w, 15);
                    GroundWiring(trans, btr, w, 25);
                }
                
                if(w.WireType == "3F+N")
                {
                    NeutralWiring(trans, btr, w, 0);
                    LiveWiring(trans, btr, w, 5);
                    LiveWiring(trans, btr, w, 10);
                    LiveWiring(trans, btr, w, 15);
                }
                if (w.WireType == "F+N")
                {
                    NeutralWiring(trans, btr, w, 0);
                    LiveWiring(trans, btr, w, 5);
                }

                if (w.WireType == "T Bus")
                {
                    GroundWiringMainBus(trans, btr, w);
                }
                if (w.WireType == "N Bus")
                {
                    NeutralWiringMainBus(trans, btr, w);
                }

            }
        }

        public void LiveWiring(Transaction trans, BlockTableRecord btr, Wiring w, double spacing)
        {
            Line Wire = new Line();
            Wire.StartPoint = new Point3d(w.StartPoint.X + spacing, w.StartPoint.Y + phaseLength/2, 0);
            Wire.EndPoint = new Point3d(w.StartPoint.X + spacing, w.StartPoint.Y - phaseLength/2, 0);
            Wire.LineWeight = w.LineWeight;
            Wire.Layer = w.LayerName.ToString();

            btr.AppendEntity(Wire);
            trans.AddNewlyCreatedDBObject(Wire, true);
        }

        public void NeutralWiring(Transaction trans, BlockTableRecord btr, Wiring w, double spacing)
        {
            Polyline Wire = new Polyline();
            Wire.AddVertexAt(0, new Point2d(w.StartPoint.X - 5 + spacing, w.StartPoint.Y + phaseLength / 2), 0, 0, 0);
            Wire.AddVertexAt(1, new Point2d(w.StartPoint.X + spacing, w.StartPoint.Y + phaseLength / 2), 0, 0, 0);
            Wire.AddVertexAt(2, new Point2d(w.StartPoint.X + spacing, w.StartPoint.Y - phaseLength / 2), 0, 0, 0);
            Wire.LineWeight = w.LineWeight;
            Wire.Layer = w.LayerName.ToString();
            
            btr.AppendEntity(Wire);
            trans.AddNewlyCreatedDBObject(Wire, true);
        }

        public void GroundWiring(Transaction trans, BlockTableRecord btr, Wiring w, double spacing)
        {
            Line Wire = new Line();
            Wire.StartPoint = new Point3d(w.StartPoint.X - 5 + spacing, w.StartPoint.Y + phaseLength / 2, 0);
            Wire.EndPoint = new Point3d(w.StartPoint.X + 5+ spacing, w.StartPoint.Y + phaseLength / 2, 0);
            Wire.LineWeight = w.LineWeight;
            Wire.Layer = w.LayerName.ToString();

            Line Wire2 = new Line();
            Wire2.StartPoint = new Point3d(w.StartPoint.X + spacing, w.StartPoint.Y + phaseLength / 2, 0);
            Wire2.EndPoint = new Point3d(w.StartPoint.X+ spacing, w.StartPoint.Y - phaseLength / 2, 0);
            Wire2.LineWeight = w.LineWeight;
            Wire2.Layer = w.LayerName.ToString();

            btr.AppendEntity(Wire);
            trans.AddNewlyCreatedDBObject(Wire, true);

            btr.AppendEntity(Wire2);
            trans.AddNewlyCreatedDBObject(Wire2, true);
        }

        public void GroundWiringMainBus(Transaction trans, BlockTableRecord btr, Wiring w)
        {
            Line Wire = new Line();
            Wire.StartPoint = new Point3d(w.StartPoint.X - phaseLength / 2 , w.StartPoint.Y - 5, 0);
            Wire.EndPoint = new Point3d(w.StartPoint.X - phaseLength / 2, w.StartPoint.Y + 5, 0);
            Wire.LineWeight = w.LineWeight;
            Wire.Layer = w.LayerName.ToString();

            Line Wire2 = new Line();
            Wire2.StartPoint = new Point3d(w.StartPoint.X + phaseLength / 2, w.StartPoint.Y, 0);
            Wire2.EndPoint = new Point3d(w.StartPoint.X - phaseLength / 2, w.StartPoint.Y, 0);
            Wire2.LineWeight = w.LineWeight;
            Wire2.Layer = w.LayerName.ToString();

            btr.AppendEntity(Wire);
            trans.AddNewlyCreatedDBObject(Wire, true);

            btr.AppendEntity(Wire2);
            trans.AddNewlyCreatedDBObject(Wire2, true);
        }

        public void NeutralWiringMainBus(Transaction trans, BlockTableRecord btr, Wiring w)
        {
            Polyline Wire = new Polyline();
            Wire.AddVertexAt(0, new Point2d(w.StartPoint.X - phaseLength / 2, w.StartPoint.Y - 5), 0, 0, 0);
            Wire.AddVertexAt(1, new Point2d(w.StartPoint.X - phaseLength / 2, w.StartPoint.Y), 0, 0, 0);
            Wire.AddVertexAt(2, new Point2d(w.StartPoint.X + phaseLength / 2, w.StartPoint.Y), 0, 0, 0);
            Wire.LineWeight = w.LineWeight;
            Wire.Layer = w.LayerName.ToString();

            btr.AppendEntity(Wire);
            trans.AddNewlyCreatedDBObject(Wire, true);
        }




    }
}
