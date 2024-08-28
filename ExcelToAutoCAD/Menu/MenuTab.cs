

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace ExcelToAutoCAD.Menu
{
    class MenuTab
    {
        private ExcelToAutoCAD _exTADCAD;
        // [CommandMethod("testmyRibbon", CommandFlags.Transparent)]
        public void OpenMenu(ExcelToAutoCAD rdExcel)
        {
            _exTADCAD = rdExcel;
                              
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon != null)
            {
                RibbonTab rtab = ribbon.FindTab("SheetToCAD");
                if (rtab != null)
                {
                    ribbon.Tabs.Remove(rtab);
                }
                rtab = new RibbonTab();
                rtab.Title = "Sheet To CAD";
                rtab.Id = "sheettocad";                
                //Add the Tab
                ribbon.Tabs.Add(rtab);
                addContent(rtab);
            }
        }

        static void addContent(RibbonTab rtab)
        {
            rtab.Panels.Add(AddOnePanel());
            
        }

        static RibbonPanel AddOnePanel()
        {
            
            RibbonButton rb;
            RibbonPanelSource rps = new RibbonPanelSource();
            rps.Title = "Menu";
            RibbonPanel rp = new RibbonPanel();
            rp.Source = rps;

            //Create a Command Item that the Dialog Launcher can use,
            // for this test it is just a place holder.
            RibbonButton rci = new RibbonButton();
            rci.Name = "SheetToCAD";
          
            //assign the Command Item to the DialgLauncher which auto-enables
            // the little button at the lower right of a Panel
            rps.DialogLauncher = rci;                                  
            
            rb = new RibbonButton();
            rb.Name = "Sheet To CAD";
            rb.ShowText = true;
            rb.Text = "Sheet To CAD";

            rb.Image = ResourceImage.imageRibbon.ToBitmapImage();

            rb.Size = RibbonItemSize.Standard;
            rb.ShowImage = true;

            rb.CommandHandler = new RibbonButtonCommandHandler();
            rb.CommandParameter = "._STC-SHEET-TO-CAD ";

            //Add the Button to the Tab
            rps.Items.Add(rb);     
            
            RibbonButton rb_help = new RibbonButton();
            rb_help.Text = "Ajuda";
            rb_help.Size = RibbonItemSize.Standard;
            rb_help.ShowText = true;
                                   

            rb_help.Image = ResourceImage.HelpIcon.ToBitmapImage();

            rps.Items.Add(rb_help);

            rb_help.CommandHandler = new RibbonButtonCommandHandler();
            rb_help.CommandParameter = "._STC-HELP ";

                        
            return rp;
        }

    }
 }
