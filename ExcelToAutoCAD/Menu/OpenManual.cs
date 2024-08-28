using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;


namespace ExcelToAutoCAD.Menu
{
    public class OpenManual
    {
        public void Open()
        {
            try
            {
                string relativeFolderPath = "Manual.pdf";

                string currentDirectory = Environment.CurrentDirectory;

                string folderPath = Path.Combine(currentDirectory, relativeFolderPath);

                if (File.Exists(folderPath))
                {
                    Process.Start(folderPath);
                }
                else
                {
                    MessageBox.Show("O arquivo não foi encontrado", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                               
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocorreu um erro ao abrir o arquivo: " + ex.Message);
            }



        }
    }

    
}
