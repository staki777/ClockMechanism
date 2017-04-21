using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace Macro2.csproj
{
    public partial class SolidWorksMacro
    {

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;

        public void Main()
        {
            string folderPath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Wybierz folder, w którym maj¹ byæ zapisane pliki.";
            while (folderPath == "")
            {
                fbd.ShowDialog();
                folderPath = fbd.SelectedPath;
            }
   
            //string folderPath = "C:\\Users\\user\\Desktop\\a";
         //0.04 zamiast 0.06
            ClockMechanismCreator cmc = new ClockMechanismCreator(0.02, 0.05, 0.08, 0.05, 0.003, 0.01 , folderPath, swApp);
            cmc.CreateMechanism();
        }


    }
}


