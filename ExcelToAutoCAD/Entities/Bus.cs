using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;


namespace ExcelToAutoCAD
{
    internal class Bus
    {
        public Point3d StartPoint{ get; set; }
        public Point3d EndPoint{ get; set; }

        Autodesk.AutoCAD.Colors.Color ColorIndex;

        public string LayerName { get; set; }

        public LineWeight LineWeight_ { get; set; }
        
        public string Linetype { get; set; }


        public List<Bus> busses { get; set; } = new List<Bus>();

        public Bus() { }             


        public Bus(Point3d startPoint, Point3d endPoint, LineWeight lineWeight, string linetype, string layerName)
        {
            StartPoint = startPoint;
            EndPoint = endPoint;            
            LineWeight_ = lineWeight;
            Linetype = linetype;
            LayerName = layerName;
        }

        public void AddBus(Point3d startPoint, Point3d endPoint, LineWeight lineWeight, string linetype, string layerName)
        {
            busses.Add(new Bus(startPoint, endPoint, lineWeight, linetype, layerName));
        }

        public void DrawBus(Transaction trans, BlockTableRecord btr)
        {                      

                foreach (Bus b in busses)
            {
                Autodesk.AutoCAD.DatabaseServices.Line busLine = new Autodesk.AutoCAD.DatabaseServices.Line(b.StartPoint, b.EndPoint);                
                busLine.LineWeight = b.LineWeight_;
                
                //string linetype = "HIDDEN";
                busLine.Linetype = b.Linetype;
                busLine.Layer = b.LayerName.ToString();
                

                btr.AppendEntity(busLine);
                trans.AddNewlyCreatedDBObject(busLine, true);
                
            }           
            
        }
    }
}
