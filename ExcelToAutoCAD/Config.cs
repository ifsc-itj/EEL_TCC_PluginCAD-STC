using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToAutoCAD
{
    public partial class Config : Form
    {
        public Config()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();

            // Exibe o diálogo de seleção de cor
            DialogResult result = colorDialog.ShowDialog();

            // Verifica se o usuário selecionou uma cor e atualiza o componente conforme necessário
            if (result == DialogResult.OK)
            {
                // Obtém a cor selecionada
                Color corSelecionada = colorDialog.Color;

                // Atualiza o componente ou realiza outras operações com a cor selecionada
                // Exemplo: definir a cor de fundo de um painel (Panel)
                //panel1.BackColor = corSelecionada;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
 }
