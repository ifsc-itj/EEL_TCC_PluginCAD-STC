using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToAutoCAD.Menu
{
    public class RibbonButtonCommandHandler : System.Windows.Input.ICommand
    {
        public bool CanExecute(object parameter)
        {
            return true; //significa que o botão sempre está ativado
        }
        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            RibbonCommandItem cmd = parameter as RibbonCommandItem;
            Document dwg = Application.DocumentManager.MdiActiveDocument;
            dwg.SendStringToExecute((string)cmd.CommandParameter, true, false, true);
        }
    }
}
