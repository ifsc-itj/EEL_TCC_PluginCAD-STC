using ExcelToAutoCAD.Save;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD.DataSaves
{
    public class DataMain
    {
        public DataMaterialList dataMaterialList{ get; set; }
        public DataInsertFile dataInsertFile { get; set; }

        public DataConfig dataConfig { get; set; }
    }
}
