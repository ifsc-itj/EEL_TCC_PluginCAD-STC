using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Core;
using System.IO;

namespace ExcelToAutoCAD
{
    
    public partial class InsertForm : Form
    {
        Point3d insPt;
        double _posX, _posY;
        string pathFile;

        public InsertForm()
        {
            InitializeComponent();
        }

        private void btOk_Click(object sender, EventArgs e)
        {
            if (gbCoordinates.Enabled == true)
            {
                bool isValid = ValidateEntry(txtCoordX);
                if (isValid == false)
                {
                    MessageBox.Show("Coordenada X com entrada inválida. Por favor, entre com um valor válido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtCoordX.Focus();
                    return;
                }
                isValid = ValidateEntry(txtCoordY);
                if (isValid == false)
                {
                    MessageBox.Show("Coordenada Y com entrada inválida. Por favor, entre com um valor válido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtCoordY.Focus();
                    return;
                }
                _posX = double.Parse(txtCoordX.Text.Trim());
                _posY = double.Parse(txtCoordY.Text.Trim());
                insPt = new Point3d(_posX, _posY, 0);
            }
            if (rbScreen.Checked)
            {
                this.Hide();
                Editor edt = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                PromptPointOptions ppo = new PromptPointOptions("Escolha o ponto inicial: ");
                PromptPointResult ppr = edt.GetPoint(ppo);
                insPt = ppr.Value;
                this.Show();
            }                       

            ExcelToAutoCAD exTACAD = new ExcelToAutoCAD();
            if(!string.IsNullOrWhiteSpace(pathFile))
            {
                exTACAD.ExcelToCAD(insPt, pathFile);
            }
            else
            {
                MessageBox.Show("Selecione um arquivo!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtFilePath.Focus();
                return;
            }            

            lbInfo.Text = "Unifilar criado com sucesso!";
            lbInfo.ForeColor = Color.Green;
        }

        private void rbScreen_CheckedChanged(object sender, EventArgs e)
        {
            gbCoordinates.Enabled = false;
        }

        private void rbCoordinates_CheckedChanged(object sender, EventArgs e)
        {
            gbCoordinates.Enabled = true;
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btFileSearch_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "(*.xlsx)|*.xlsx";
            if(dlg.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = dlg.FileName;
                pathFile = dlg.FileName;
            }
        }


        private bool ValidateEntry(TextBox tb) //método para validar entrada de dados nas coordenadas
        {
            bool isValid = false;
            double value;

            try
            {
                if(tb.Text.Trim() == "")
                {
                    lbInfo.Text = "Por favor, entre com um valor.";
                    lbInfo.ForeColor = Color.Red;
                }
                else
                {
                    value = double.Parse(tb.Text.Trim());
                    isValid = true;
                }
            }
            catch (Exception ex)
            {
                lbInfo.Text = "Entrada inválida: " + ex.Message;
                lbInfo.ForeColor = Color.Red;
                
            }

            return isValid;
        }

    }
}
