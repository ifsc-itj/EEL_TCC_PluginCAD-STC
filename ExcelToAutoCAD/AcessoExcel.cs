using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Security.Cryptography.X509Certificates;
using System.Drawing.Printing;

namespace ExcelToAutoCAD
{
    public class AcessoExcel
    {
        public enum Fase { A, B, C }
        Fase fase;
        
        public struct Row
        {
            //DESCRIÇÃO
            public string Circuito; //coluna 1
            public string Descricao; //coluna 2
            //POTENCIAS
            public string Potencia; //coluna 3
            public string FaseA; //coluna 6
            public string FaseB;//couluna 7 
            public string FaseC; //couluna 8 
            public Fase Fase; //coluna 6,7 e 8
            public string CorrenteCircuitos; //coluna 9

            //CABOS
            public string QuantidadeCabos; //coluna 10
            public string DividirCabos; //coluna 11
            public string ReduzirTerra;//coluna 12
            public string SecaoCondutor;//coluna 13
            public string TipoIsolacao;//coluna 14
            //DISJUNTORES
            public string TipoDisjuntor;//coluna 15
            public string CurvaDisjuntor;//coluna 16
            public string QuantPolosDisjuntor;//coluna 17
            public string CorrenteDisjuntor;//coluna 18
            public string CorrenteCCDisjuntor;//coluna 19
            //DRs
            public string NumeracaoIDR;//coluna 20
            public string QuantPolosIDR;//coluna 21
            public string CorrenteIDR;//coluna 22
            public string FugaIDR;//coluna 23



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
                       object fugaIDR) : this()
            {
                if (circuito == null) circuito = "";
                Circuito = circuito.ToString();
                if (descricao == null) descricao = "";
                Descricao = descricao.ToString();
                if (potencia == null) potencia = "";
                Potencia = potencia.ToString();
                if (faseA == null) faseA = "";
                if (faseB == null) faseB = "";
                if (faseC == null) faseC = "";
                if (correnteCircuitos == null) correnteCircuitos = "";
                CorrenteCircuitos = correnteCircuitos.ToString();
                if(quantidadeCabos == null) quantidadeCabos="";
                QuantidadeCabos = quantidadeCabos.ToString();
                if (dividirCabos == null) dividirCabos = "";
                DividirCabos = dividirCabos.ToString();
                if (reduzirTerra == null) reduzirTerra = "";
                ReduzirTerra = reduzirTerra.ToString();
                if (secaoCondutor == null) secaoCondutor = "";
                SecaoCondutor = secaoCondutor.ToString();
                if (tipoIsolacao == null) tipoIsolacao = "";
                TipoIsolacao = tipoIsolacao.ToString();
                if (tipoDisjuntor == null) tipoDisjuntor = "";
                TipoDisjuntor = tipoDisjuntor.ToString();
                if (curvaDisjuntor == null) curvaDisjuntor = "";
                CurvaDisjuntor = curvaDisjuntor.ToString();
                if (quantPolosDisjuntor == null) quantPolosDisjuntor = "";
                QuantPolosDisjuntor = quantPolosDisjuntor.ToString();
                if (correnteDisjuntor == null) correnteDisjuntor = "";
                CorrenteDisjuntor = correnteDisjuntor.ToString();
                if (correnteCCDisjuntor == null) correnteCCDisjuntor = "";
                CorrenteCCDisjuntor = correnteCCDisjuntor.ToString();
                if (numeracaoIDR == null) numeracaoIDR = "";
                NumeracaoIDR = numeracaoIDR.ToString();
                if (quantPolosIDR == null) quantPolosIDR = "";
                QuantPolosIDR = quantPolosIDR.ToString();
                if (correnteIDR == null) correnteIDR = "";
                CorrenteIDR = correnteIDR.ToString();
                if (fugaIDR == null) fugaIDR = "";
                FugaIDR = fugaIDR.ToString();
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
                       object correnteIDR)  : this()
            {
                if (circuito == null) circuito = "";
                Circuito = circuito.ToString();
                if (descricao == null) descricao = "";
                Descricao = descricao.ToString();
                if (potencia == null) potencia = "";
                Potencia = potencia.ToString();
                if (correnteCircuitos == null) correnteCircuitos = "";
                CorrenteCircuitos = correnteCircuitos.ToString();
                if (quantidadeCabos == null) quantidadeCabos = "";
                QuantidadeCabos = quantidadeCabos.ToString();
                if (dividirCabos == null) dividirCabos = "";
                DividirCabos = dividirCabos.ToString();
                if (reduzirTerra == null) reduzirTerra = "";
                ReduzirTerra = reduzirTerra.ToString();
                if (secaoCondutor == null) secaoCondutor = "";
                SecaoCondutor = secaoCondutor.ToString();
                if (tipoDisjuntor == null) tipoDisjuntor = "";
                TipoDisjuntor = tipoDisjuntor.ToString();
                if (curvaDisjuntor == null) curvaDisjuntor = "";
                CurvaDisjuntor = curvaDisjuntor.ToString();
                if (quantPolosDisjuntor == null) quantPolosDisjuntor = "";
                QuantPolosDisjuntor = quantPolosDisjuntor.ToString();
                if (correnteDisjuntor == null) correnteDisjuntor = "";
                CorrenteDisjuntor = correnteDisjuntor.ToString();
                if (numeracaoIDR == null) numeracaoIDR = "";
                NumeracaoIDR = numeracaoIDR.ToString();
                if (quantPolosIDR == null) quantPolosIDR = "";
                QuantPolosIDR = quantPolosIDR.ToString();
                if (correnteIDR == null) correnteIDR = "";
                CorrenteIDR = correnteIDR.ToString();
            }

        }

        public List<Row> rows = new List<Row>();

        private int _linhaInicial = 4; //linha incial de leitura no excel

        private int[] _coluna = Enumerable.Range(0,31).ToArray(); //cria uma vetor para acessar as colunas do excel

        private string PromptFile() //// ********** ATENÇÃO:  desativar metodo **************
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            Editor edt = doc.Editor;

            PromptOpenFileOptions openFileOpts = new PromptOpenFileOptions("Selecione o arquivo excel: ");
            openFileOpts.Filter = "Arquivos do Excel (*.xlsx)|*.xlsx";
            PromptFileNameResult openFileRes = edt.GetFileNameForOpen(openFileOpts);
            if (openFileRes.Status != PromptStatus.OK)
                return null;
            return openFileRes.StringResult;
        }

        private string CopyFile(string fileName)
        {
            string tempFile = Path.GetTempFileName();
            File.Copy(fileName, tempFile, true);
            return tempFile;
        }

        public void LerColuna(string pathFile)
        {          
            string fileName = pathFile;
            string tempFile = CopyFile(fileName);

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(tempFile);

            // seleciona a planilha
            Worksheet worksheet = workbook.Sheets["QD"];

            // lê os valores da coluna até encontrar uma célula em branco
            bool continuaLendo = true;
            while (continuaLendo)
            {
                // lê o valor da célula atual
                Range rangeCircuito = worksheet.Cells[_linhaInicial, _coluna[1]]; //circuito
                Range rangeDescricao = worksheet.Cells[_linhaInicial, _coluna[2]]; //descricao      
                Range rangePotencia = worksheet.Cells[_linhaInicial, _coluna[3]]; //potencia
                Range rangeFaseA = worksheet.Cells[_linhaInicial, _coluna[6]]; //fase A
                Range rangeFaseB = worksheet.Cells[_linhaInicial, _coluna[7]]; //fase B
                Range rangeFaseC = worksheet.Cells[_linhaInicial, _coluna[8]]; //fase A
                Range rangeCorrenteCircuitos = worksheet.Cells[_linhaInicial, _coluna[9]]; //correnteCircuitos
                Range rangeQuantidadeCabos = worksheet.Cells[_linhaInicial, _coluna[10]]; //quantidadeCabos
                Range rangeDividirCabos = worksheet.Cells[_linhaInicial, _coluna[11]]; //dividirCabos
                Range rangeReduzirTerra = worksheet.Cells[_linhaInicial, _coluna[12]]; //reduzirTerra
                Range rangeSecaoCondutor = worksheet.Cells[_linhaInicial, _coluna[13]]; //secaoCondutor
                Range rangeTipoIsolacao = worksheet.Cells[_linhaInicial, _coluna[14]]; //tipoIsolacao
                Range rangeTipoDisjuntor = worksheet.Cells[_linhaInicial, _coluna[15]]; // TipoDisjuntor
                Range rangeCurvaDisjuntor = worksheet.Cells[_linhaInicial, _coluna[16]]; // CurvaDisjuntor
                Range rangeQuantPolosDisjuntor = worksheet.Cells[_linhaInicial, _coluna[17]]; // QuantPolosDisjuntor
                Range rangeCorrenteDisjuntor = worksheet.Cells[_linhaInicial, _coluna[18]]; // CorrenteDisjuntor
                Range rangeCorrenteCCDisjuntor = worksheet.Cells[_linhaInicial, _coluna[19]]; // CorrenteCCDisjuntor
                Range rangeNumeracaoIDR = worksheet.Cells[_linhaInicial, _coluna[20]]; // NumeracaoIDR
                Range rangeQuantPolosIDR = worksheet.Cells[_linhaInicial, _coluna[21]]; // QuantPolosIDR
                Range rangeCorrenteIDR = worksheet.Cells[_linhaInicial, _coluna[22]]; // CorrenteIDR
                Range rangeFugaIDR = worksheet.Cells[_linhaInicial, _coluna[23]]; // FugaIDR
              

                if (rangeFaseA.Value != null)
                {
                    fase = Fase.A;
                }
                else if (rangeFaseB.Value != null)
                {
                    fase = Fase.B;
                }
                else if (rangeFaseC.Value != null)
                {
                    fase = Fase.C;
                }
                else
                {
                    fase = 0;
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
                                    rangeFugaIDR.Value
                                );
                //Row row = new Row(rangeCircuito.Value, rangeDescricao.Value, rangeCondutor.Value, rangeDisjuntor.Value);
                rows.Add(row);

                // verifica se a célula está em branco
                if (string.IsNullOrEmpty(row.Descricao))
                    continuaLendo = false;
                else
                    _linhaInicial++;
            }


        workbook.Close();
            excelApp.Quit();
            File.Delete(tempFile);
        }
        

    }
}
