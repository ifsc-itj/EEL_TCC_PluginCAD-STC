using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Navigation;
using static ExcelToAutoCAD.ExcelAccess;

namespace ExcelToAutoCAD
{
    public class ListaMaterial
    {        

        private Dictionary<string, int> breakersCount = new Dictionary<string, int>(); //contabilização de itens

        private Dictionary<string, int> terminalsCount = new Dictionary<string, int>();

        private Dictionary<string, int> RDCCount = new Dictionary<string, int>();

        private Dictionary<string, int> SPDCount = new Dictionary<string, int>();

        int totalDINspace = 0;
        string dps = "";
        string breaker = "";
        string terminal = "";
        string terminalType = "TERMINAL TUBULAR ";
        string RDC = "";
        string RDCType = "DR ";
        string previusRow = "-1";
        ExcelAccess ReadExcel = new ExcelAccess();

        public ListaMaterial(string pathFile, string sheetName) {
                        
            ReadExcel.ReadColumn(pathFile, sheetName);                 

            ExcelAccess.Row row = new ExcelAccess.Row();                      

            for (int i = 0; i < ReadExcel.rows.Count; i++)
            {                
                row = ReadExcel.rows[i];
                //CONTABILIZA DISJUNTOR
                if (!string.IsNullOrEmpty(row.CBCurrent))
                {
                    breaker = "DISJUNTOR " + row.CBCurve + " " + row.CBCurrent.ToString() + " " + row.CBPolesQuantity;
                }

                if(breakersCount.ContainsKey(breaker) ) {
                    breakersCount[breaker]++;  
                    totalDINspace++;
                }
                else if(breaker != "")
                {
                    breakersCount[breaker] = 1;
                    totalDINspace++;
                }
                breaker = "";

                TerminalCount(row);

                //CONTABILIZA DR
                if (row.RCDNumbering != "0" && row.RCDNumbering != previusRow)
                {
                    RDC = row.RCDPolesQuant + " " + row.RCDCurrent + " Fuga " + row.RCDLeakage;
                }
                if (RDCCount.ContainsKey(RDC))
                {
                    RDCCount[RDC]++;
                    if (row.RCDPolesQuant == "BIPOLAR")
                        totalDINspace += 2;
                    else if (row.RCDPolesQuant == "TETRAPOLAR")
                        totalDINspace += 4;
                }
                else if(RDC != "" && !string.IsNullOrEmpty(row.RCDNumbering))
                {
                    RDCCount[RDC] = 1;
                    if (row.RCDPolesQuant == "BIPOLAR")
                        totalDINspace += 2;
                    else if (row.RCDPolesQuant == "TETRAPOLAR")
                        totalDINspace += 4;
                }
                RDC = "";
                previusRow = row.RCDNumbering;
            }

            for (int i = 0; i < ReadExcel.others.Count; i++)
            {
                row = ReadExcel.others[i];

                if (row.Circuit == "DPS" && !string.IsNullOrEmpty(row.Description))
                {
                    if(row.CBPolesQuantity == "TRIPOLAR")
                    {
                        SPDCount[row.Description] = 4; //soma com o neutro
                        totalDINspace += 4;
                    }
                    else if((row.CBPolesQuantity == "BIPOLAR"))
                    {
                        SPDCount[row.Description] = 3;
                        totalDINspace += 3;
                    }
                    else
                    {
                        SPDCount[row.Description] = 2;
                        totalDINspace += 2;
                    }
                    TerminalCount(row);
                }
                if (row.Circuit == "PROTEÇÃO DPS" && !string.IsNullOrEmpty(row.Description))
                {
                    breaker = "DISJUNTOR " + row.CBCurve + " " + row.CBCurrent.ToString() + " " + row.CBPolesQuantity;
                    if (breakersCount.ContainsKey(breaker))
                    {
                        breakersCount[breaker]++;
                        if (row.CBType == "DIN")
                        {
                            if (row.CBPolesQuantity == "TRIPOLAR")
                                totalDINspace += 3;
                            else if (row.CBPolesQuantity == "BIPOLAR")
                                totalDINspace += 2;
                            else
                                totalDINspace++;
                        }
                    }
                    else if (breaker != "")
                    {
                        breakersCount[breaker] = 1;
                        if (row.CBType == "DIN")
                            totalDINspace++;
                    }
                    TerminalCount(row);
                }
                if (row.Circuit == "ENTRADA DE ENERGIA")
                {
                    breaker = "DISJUNTOR " + row.CBCurve + " " + row.CBCurrent.ToString() + " " + row.CBPolesQuantity;
                    if (breakersCount.ContainsKey(breaker))
                    {
                        breakersCount[breaker]++;
                        if(row.CBType == "DIN")
                        {
                            if (row.CBPolesQuantity == "TRIPOLAR")
                                totalDINspace += 3;
                            else if (row.CBPolesQuantity == "BIPOLAR")
                                totalDINspace += 2;
                            else
                                totalDINspace++;
                        }                          
                           
                    }
                    else if (breaker != "")
                    {
                        breakersCount[breaker] = 1;
                        if (row.CBType == "DIN")
                            if (row.CBPolesQuantity == "TRIPOLAR")
                                totalDINspace += 3;
                            else if (row.CBPolesQuantity == "BIPOLAR")
                                totalDINspace += 2;
                            else
                                totalDINspace++;
                    }
                    TerminalCount(row);
                }
            }

            }

        public void TerminalCount(ExcelAccess.Row row)
        {
            
            if (!string.IsNullOrEmpty(row.ConductorSection))
            {
                terminal = terminalType + row.ConductorSection;
            }

            if (terminalsCount.ContainsKey(terminal))
            {
                if (row.CBPolesQuantity == "UNIPOLAR")
                    terminalsCount[terminal]++;
                else if (row.CBPolesQuantity == "TRIPOLAR")
                    terminalsCount[terminal] += 3;
                else
                    terminalsCount[terminal] += 2;
            }
            else
            {
                terminalsCount[terminal] = 1;
            }

        }

        public Dictionary<string, int> GetBreakersCount(){ return breakersCount; }

        public Dictionary<string, int> GetTerminalsCount() { return terminalsCount; }

        public Dictionary<string, int> GetRCDCount() { return RDCCount; }

        public Dictionary<string, int> GetSPD() { return SPDCount; }
                

        public int GetPanelSize() { return totalDINspace; }
    }
}


