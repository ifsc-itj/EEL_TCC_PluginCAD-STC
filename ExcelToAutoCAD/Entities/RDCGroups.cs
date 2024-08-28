using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD.Entities
{
    internal class RDCGroups
    {
        public string RDCNumber { get; set; }

        public int RDCCircuitsGroup { get; set; }
        public string RDCPoles { get; set; }

        public string RDCCurrent { get; set;}

        public string RDCLeakage { get; set; }

        public string RDCClass { get; set; }

        public string RDCCableSection { get; set; }

        List<RDCGroups> RDCs = new List<RDCGroups>();              

        public RDCGroups() { }

        public RDCGroups(string rDCNumber, string rDCPoles, string rDCCurrent, string rCDLeakage, string rCDClass, string rDCCableSection, int rDCCircuitsGroup)
        {
            RDCNumber = rDCNumber;
            RDCPoles = rDCPoles;
            RDCCurrent = rDCCurrent;
            RDCLeakage = rCDLeakage;
            RDCClass = rCDClass;
            RDCCableSection = rDCCableSection;
            RDCCircuitsGroup = rDCCircuitsGroup;
        }        

        public void AddRDC (string rDCNumber, string rDCPoles, string rDCCurrent, string rCDLeakage, string rCDClass, string rDCCableSection, int rDCCircuitsGroup)
        {
            RDCs.Add(new RDCGroups(rDCNumber, rDCPoles, rDCCurrent, rCDLeakage, rCDClass, rDCCableSection, rDCCircuitsGroup));
        }
              

        public RDCGroups ReadRDC(int position)
        {
            if (position >= 0)
            {
                RDCGroups R = RDCs[position];
                return R;
            } else {
                throw new ArgumentOutOfRangeException("position", "Posição fora dos limites da lista");
            }
        }          


    }
}
