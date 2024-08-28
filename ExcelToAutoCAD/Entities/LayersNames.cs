using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD.Entities
{
    public class LayersNames
    {
        public string LayerName { get; set; }

        public int ColorIndex { get; set; }

        List<LayersNames> Layers { get; set; } = new List<LayersNames>();

        public LayersNames() { }

        public LayersNames(string layerName, int colorIndex) { 
            
            LayerName = layerName;
            ColorIndex = colorIndex;
        }
        public void AddLayer(string layerName, int colorIndex)
        {
            Layers.Add(new LayersNames(layerName, colorIndex));
        }
                
        public List<LayersNames> GetLayers() { return Layers; }

    }
}
