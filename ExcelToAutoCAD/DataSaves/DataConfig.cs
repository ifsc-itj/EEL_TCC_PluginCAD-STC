using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD.DataSaves
{
    public class DataConfig
    {

        public string dataScale { get; set; }
        public string dataLCircuits { get; set; }
        public string dataLBus { get; set; }
        public string dataLNeutral { get; set; }
        public string dataLGround { get; set; }
        public string dataLWiring { get; set; }
        public string dataTType { get; set; }
        public string dataTSize { get; set; }

        public string dataPhaseA { get; set; }
        public string dataPhaseB { get; set;}
        public string dataPhaseC { get; set;}

    }
}
