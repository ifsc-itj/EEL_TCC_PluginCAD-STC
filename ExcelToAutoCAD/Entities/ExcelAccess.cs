using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelToAutoCAD
{
    public class ExcelAccess
    {
        public enum Phase { A, B, C }
        Phase phase;

        public struct Row
        {
            //DESCRIÇÃO
            public string Circuit; //coluna 1
            public string Description; //coluna 2
            //POTENCIAS
            public string Power; //coluna 3
            public string PhaseA; //coluna 6
            public string PhaseB;//couluna 7 
            public string PhaseC; //couluna 8 
            public Phase Phase; //coluna 6,7 e 8
            public string CircuitsCurrent; //coluna 9

            //CABOS
            public string CablesQuantity; //coluna 10
            public string SplitCables; //coluna 11
            public string MinimizeEarth;//coluna 12
            public string ConductorSection;//coluna 13
            public string IsolationType;//coluna 14
            //DISJUNTORES
            public string CBType;//coluna 15 ----    CB - Circuit Breaker
            public string CBCurve;//coluna 16
            public string CBPolesQuantity;//coluna 17
            public string CBCurrent;//coluna 18
            public string ShortCBCurrent;//coluna 19
            //DRs
            public string RCDNumbering;//coluna 20 ----        RCD - Residual Current Device
            public string RCDPolesQuant;//coluna 21
            public string RCDCurrent;//coluna 22
            public string RCDLeakage;//coluna 23

            public string Title;



            public Row(object circuito,
                       object descricao,
                       object potencia,
                       object faseA,
                       object faseB,
                       object faseC,
                       object correnteCircuitos,
                       object quantidadeCabos,
                       object dividirCabos,
                       object reduzirTerra,
                       object secaoCondutor,
                       object tipoIsolacao,
                       object tipoDisjuntor,
                       object curvaDisjuntor,
                       object quantPolosDisjuntor,
                       object correnteDisjuntor,
                       object correnteCCDisjuntor,
                       object numeracaoIDR,
                       object quantPolosIDR,
                       object correnteIDR,
                       object fugaIDR,
                       object title) : this()
            {
                if (circuito == null) circuito = "";
                Circuit = circuito.ToString();
                if (descricao == null) descricao = "";
                Description = descricao.ToString();
                if (potencia == null) potencia = "";
                Power = potencia.ToString();
                if (faseA == null) faseA = "";
                PhaseA = faseA.ToString();
                if (faseB == null) faseB = "";
                PhaseB = faseB.ToString();
                if (faseC == null) faseC = "";
                PhaseC = faseC.ToString();
                if (correnteCircuitos == null) correnteCircuitos = "";
                CircuitsCurrent = correnteCircuitos.ToString();
                if (quantidadeCabos == null) quantidadeCabos = "";
                CablesQuantity = quantidadeCabos.ToString();
                if (dividirCabos == null) dividirCabos = "";
                SplitCables = dividirCabos.ToString();
                if (reduzirTerra == null) reduzirTerra = "";
                MinimizeEarth = reduzirTerra.ToString();
                if (secaoCondutor == null) secaoCondutor = "";
                ConductorSection = secaoCondutor.ToString();
                if (tipoIsolacao == null) tipoIsolacao = "";
                IsolationType = tipoIsolacao.ToString();
                if (tipoDisjuntor == null) tipoDisjuntor = "";
                CBType = tipoDisjuntor.ToString();
                if (curvaDisjuntor == null) curvaDisjuntor = "";
                CBCurve = curvaDisjuntor.ToString();
                if (quantPolosDisjuntor == null) quantPolosDisjuntor = "";
                CBPolesQuantity = quantPolosDisjuntor.ToString();
                if (correnteDisjuntor == null) correnteDisjuntor = "";
                CBCurrent = correnteDisjuntor.ToString();
                if (correnteCCDisjuntor == null) correnteCCDisjuntor = "";
                ShortCBCurrent = correnteCCDisjuntor.ToString();
                if (numeracaoIDR == null) numeracaoIDR = "";
                RCDNumbering = numeracaoIDR.ToString();
                if (quantPolosIDR == null) quantPolosIDR = "";
                RCDPolesQuant = quantPolosIDR.ToString();
                if (correnteIDR == null) correnteIDR = "";
                RCDCurrent = correnteIDR.ToString();
                if (fugaIDR == null) fugaIDR = "";
                RCDLeakage = fugaIDR.ToString();
                if (title == null) title= "";
                Title = title.ToString();
            }


            public Row(object circuito,
                       object descricao,
                       object potencia,
                       object correnteCircuitos,
                       object quantidadeCabos,
                       object dividirCabos,
                       object reduzirTerra,
                       object secaoCondutor,
                       object tipoDisjuntor,
                       object curvaDisjuntor,
                       object quantPolosDisjuntor,
                       object correnteDisjuntor,
                       object numeracaoIDR,
                       object quantPolosIDR,
                       object correnteIDR) : this()
            {
                if (circuito == null) circuito = "";
                Circuit = circuito.ToString();
                if (descricao == null) descricao = "";
                Description = descricao.ToString();
                if (potencia == null) potencia = "";
                Power = potencia.ToString();
                if (correnteCircuitos == null) correnteCircuitos = "";
                CircuitsCurrent = correnteCircuitos.ToString();
                if (quantidadeCabos == null) quantidadeCabos = "";
                CablesQuantity = quantidadeCabos.ToString();
                if (dividirCabos == null) dividirCabos = "";
                SplitCables = dividirCabos.ToString();
                if (reduzirTerra == null) reduzirTerra = "";
                MinimizeEarth = reduzirTerra.ToString();
                if (secaoCondutor == null) secaoCondutor = "";
                ConductorSection = secaoCondutor.ToString();
                if (tipoDisjuntor == null) tipoDisjuntor = "";
                CBType = tipoDisjuntor.ToString();
                if (curvaDisjuntor == null) curvaDisjuntor = "";
                CBCurve = curvaDisjuntor.ToString();
                if (quantPolosDisjuntor == null) quantPolosDisjuntor = "";
                CBPolesQuantity = quantPolosDisjuntor.ToString();
                if (correnteDisjuntor == null) correnteDisjuntor = "";
                CBCurrent = correnteDisjuntor.ToString();
                if (numeracaoIDR == null) numeracaoIDR = "";
                RCDNumbering = numeracaoIDR.ToString();
                if (quantPolosIDR == null) quantPolosIDR = "";
                RCDPolesQuant = quantPolosIDR.ToString();
                if (correnteIDR == null) correnteIDR = "";
                RCDCurrent = correnteIDR.ToString();
            }

        }

        public List<Row> rows = new List<Row>();

        public List<Row> others = new List<Row>();

        private int firstLine = 4; //linha incial de leitura no excel

        private int[] column = Enumerable.Range(0, 31).ToArray(); //cria uma vetor para acessar as colunas do excel


        private string CopyFile(string fileName)
        {
            string tempFile = Path.GetTempFileName();
            File.Copy(fileName, tempFile, true);
            return tempFile;
        }

        public void ReadColumn(string pathFile, string sheetName)
        {
            try
            {                     

            string fileName = pathFile;
            string tempFile = CopyFile(fileName);

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(tempFile);

            // seleciona a planilha
            Worksheet worksheet = workbook.Sheets[sheetName];

            // lê os valores da coluna até encontrar uma célula em branco
            bool continuaLendo = true;

            Range rangeTitle = worksheet.Cells[1, column[2]]; //titulo do quadro
            

            while (continuaLendo)
            {
                // lê o valor da célula atual
                Range rangeCircuito = worksheet.Cells[firstLine, column[1]]; //circuito
                Range rangeDescricao = worksheet.Cells[firstLine, column[2]]; //descricao      
                Range rangePotencia = worksheet.Cells[firstLine, column[3]]; //potencia
                Range rangeFaseA = worksheet.Cells[firstLine, column[6]]; //fase A
                Range rangeFaseB = worksheet.Cells[firstLine, column[7]]; //fase B
                Range rangeFaseC = worksheet.Cells[firstLine, column[8]]; //fase A
                Range rangeCorrenteCircuitos = worksheet.Cells[firstLine, column[9]]; //correnteCircuitos
                Range rangeQuantidadeCabos = worksheet.Cells[firstLine, column[10]]; //quantidadeCabos
                Range rangeDividirCabos = worksheet.Cells[firstLine, column[11]]; //dividirCabos
                Range rangeReduzirTerra = worksheet.Cells[firstLine, column[12]]; //reduzirTerra
                Range rangeSecaoCondutor = worksheet.Cells[firstLine, column[13]]; //secaoCondutor
                Range rangeTipoIsolacao = worksheet.Cells[firstLine, column[14]]; //tipoIsolacao
                Range rangeTipoDisjuntor = worksheet.Cells[firstLine, column[15]]; // TipoDisjuntor
                Range rangeCurvaDisjuntor = worksheet.Cells[firstLine, column[16]]; // CurvaDisjuntor
                Range rangeQuantPolosDisjuntor = worksheet.Cells[firstLine, column[17]]; // QuantPolosDisjuntor
                Range rangeCorrenteDisjuntor = worksheet.Cells[firstLine, column[18]]; // CorrenteDisjuntor
                Range rangeCorrenteCCDisjuntor = worksheet.Cells[firstLine, column[19]]; // CorrenteCCDisjuntor
                Range rangeNumeracaoIDR = worksheet.Cells[firstLine, column[20]]; // NumeracaoIDR
                Range rangeQuantPolosIDR = worksheet.Cells[firstLine, column[21]]; // QuantPolosIDR
                Range rangeCorrenteIDR = worksheet.Cells[firstLine, column[22]]; // CorrenteIDR
                Range rangeFugaIDR = worksheet.Cells[firstLine, column[23]]; // FugaIDR


                if (rangeFaseA.Value != null)
                {
                    phase = Phase.A;
                }
                else if (rangeFaseB.Value != null)
                {
                    phase = Phase.B;
                }
                else if (rangeFaseC.Value != null)
                {
                    phase = Phase.C;
                }
                else
                {
                    phase = 0;
                }

                Row row = new Row(
                                    rangeCircuito.Value,
                                    rangeDescricao.Value,
                                    rangePotencia.Value,
                                    rangeFaseA.Value,
                                    rangeFaseB.Value,
                                    rangeFaseC.Value,
                                    rangeCorrenteCircuitos.Value,
                                    rangeQuantidadeCabos.Value,
                                    rangeDividirCabos.Value,
                                    rangeReduzirTerra.Value,
                                    rangeSecaoCondutor.Value,
                                    rangeTipoIsolacao.Value,
                                    rangeTipoDisjuntor.Value,
                                    rangeCurvaDisjuntor.Value,
                                    rangeQuantPolosDisjuntor.Value,
                                    rangeCorrenteDisjuntor.Value,
                                    rangeCorrenteCCDisjuntor.Value,
                                    rangeNumeracaoIDR.Value,
                                    rangeQuantPolosIDR.Value,
                                    rangeCorrenteIDR.Value,
                                    rangeFugaIDR.Value,
                                    rangeTitle.Value
                                );
  
                rows.Add(row);

                // verifica se a célula está em branco
                if (string.IsNullOrEmpty(row.Description))
                {
                    firstLine++;
                    while (continuaLendo)
                    {

                        rangeCircuito = worksheet.Cells[firstLine, column[1]]; //circuito
                        rangeDescricao = worksheet.Cells[firstLine, column[2]]; //descricao      
                        rangePotencia = worksheet.Cells[firstLine, column[3]]; //potencia
                        rangeFaseA = worksheet.Cells[firstLine, column[6]]; //fase A
                        rangeFaseB = worksheet.Cells[firstLine, column[7]]; //fase B
                        rangeFaseC = worksheet.Cells[firstLine, column[8]]; //fase A
                        rangeCorrenteCircuitos = worksheet.Cells[firstLine, column[9]]; //correnteCircuitos
                        rangeQuantidadeCabos = worksheet.Cells[firstLine, column[10]]; //quantidadeCabos
                        rangeDividirCabos = worksheet.Cells[firstLine, column[11]]; //dividirCabos
                        rangeReduzirTerra = worksheet.Cells[firstLine, column[12]]; //reduzirTerra
                        rangeSecaoCondutor = worksheet.Cells[firstLine, column[13]]; //secaoCondutor
                        rangeTipoIsolacao = worksheet.Cells[firstLine, column[14]]; //tipoIsolacao
                        rangeTipoDisjuntor = worksheet.Cells[firstLine, column[15]]; // TipoDisjuntor
                        rangeCurvaDisjuntor = worksheet.Cells[firstLine, column[16]]; // CurvaDisjuntor
                        rangeQuantPolosDisjuntor = worksheet.Cells[firstLine, column[17]]; // QuantPolosDisjuntor
                        rangeCorrenteDisjuntor = worksheet.Cells[firstLine, column[18]]; // CorrenteDisjuntor
                        rangeCorrenteCCDisjuntor = worksheet.Cells[firstLine, column[19]]; // CorrenteCCDisjuntor
                        rangeNumeracaoIDR = worksheet.Cells[firstLine, column[20]]; // NumeracaoIDR
                        rangeQuantPolosIDR = worksheet.Cells[firstLine, column[21]]; // QuantPolosIDR
                        rangeCorrenteIDR = worksheet.Cells[firstLine, column[22]]; // CorrenteIDR
                        rangeFugaIDR = worksheet.Cells[firstLine, column[23]]; // FugaIDR

                        Row other = new Row(
                                    rangeCircuito.Value,
                                    rangeDescricao.Value,
                                    rangePotencia.Value,
                                    rangeFaseA.Value,
                                    rangeFaseB.Value,
                                    rangeFaseC.Value,
                                    rangeCorrenteCircuitos.Value,
                                    rangeQuantidadeCabos.Value,
                                    rangeDividirCabos.Value,
                                    rangeReduzirTerra.Value,
                                    rangeSecaoCondutor.Value,
                                    rangeTipoIsolacao.Value,
                                    rangeTipoDisjuntor.Value,
                                    rangeCurvaDisjuntor.Value,
                                    rangeQuantPolosDisjuntor.Value,
                                    rangeCorrenteDisjuntor.Value,
                                    rangeCorrenteCCDisjuntor.Value,
                                    rangeNumeracaoIDR.Value,
                                    rangeQuantPolosIDR.Value,
                                    rangeCorrenteIDR.Value,
                                    rangeFugaIDR.Value,
                                    rangeTitle.Value
                                );

                        others.Add(other);
                        if (string.IsNullOrEmpty(other.Circuit))
                        {
                            continuaLendo = false;
                        }
                        else firstLine++;
                    }
                }

                else
                    firstLine++;
            }


            workbook.Close();
            excelApp.Quit();
            File.Delete(tempFile);
        }
            catch (System.Exception ex)
            {
                MessageBox.Show("Selecione um arquivo válido!\n" + "Erro: " + ex, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
                
            }
        }

        public List<string> GetExcelHeaders(string pathFile, string sheetName)
        {
            List<string> headers = new List<string>();

            try
            {
                string fileName = pathFile;
                string tempFile = CopyFile(fileName);

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(tempFile);

                // Seleciona a planilha
                Worksheet worksheet = workbook.Sheets[sheetName];

                // Obtém os cabeçalhos da primeira linha
                for (int i = 0; i < column.Length; i++)
                {
                    Range headerCell = worksheet.Cells[3, column[i]];
                    if(headerCell != null && headerCell.Value != null){
                        string headerValue = (headerCell.Value != null) ? headerCell.Value.ToString() : "";
                        headers.Add(headerValue);
                    }

                    
                }

                workbook.Close();
                excelApp.Quit();
                File.Delete(tempFile);

                return headers;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Selecione um arquivo válido!\n" + "Erro: " + ex, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }


    }
}
