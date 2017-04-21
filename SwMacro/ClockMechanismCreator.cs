using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Threading;

namespace Macro2.csproj
{
    public class ClockMechanismCreator
    {
        string folderPath;
        string assemblyName = "clockMechanism";
        double distanceFactor;
        SldWorks swApp;
        List<MechanismUnit> mechanismUnits = new List<MechanismUnit>();
        double r1, r2, h1, h2, stickR, wheelWidth;
        List<Component2> components = new List<Component2>();
        ModelDoc2 swDoc;
       

        public ClockMechanismCreator(double _r1, double _r2, double _h1, double _h2, double _stickR, double _wheelWidth, string _folderPath, SldWorks _swApp)
        {
            folderPath = _folderPath;
            swApp = _swApp;
            r1 = _r1;
            r2 = _r2;
            h1 = _h1;
            h2 = _h2;
            distanceFactor = (r1 + r2) / 0.07 * 0.003;
            stickR = _stickR;
            wheelWidth = _wheelWidth;
            swDoc = ((ModelDoc2)(swApp.NewAssembly()));
            
            //wstêpny zapis
            swDoc.SaveAs(folderPath + "\\" + assemblyName + ".SLDASM");
        }

        public void CreateMechanism()
        {
            mechanismUnits.Add(new MechanismUnit(r1, r2, true, wheelWidth, stickR, 0, 0, 0, h1, h2, folderPath, 1, swApp));
            mechanismUnits.Add(new MechanismUnit(r1, r2, true, wheelWidth, stickR, r1 + r2-distanceFactor, h1-h2, 0, h1, h2, folderPath, 2, swApp));
            mechanismUnits.Add(new MechanismUnit(r1, r2, false, wheelWidth, stickR, 2*(r1 + r2 - distanceFactor), 2*(h1-h2), 0, h1, h2, folderPath, 3, swApp));
            foreach (MechanismUnit me in mechanismUnits)
                me.makeMechanismUnit(swApp);

            makeMechanism();
        }

        private void makeMechanism()
        {
            bool boolstatus;
            SketchPoint skPoint = null;
            int longstatus, errors = 0;
        
            swApp.ActivateDoc3(folderPath + "\\" + assemblyName + ".SLDASM", true, (int)swRebuildOnActivation_e.swUserDecision, ref errors);
            swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            AssemblyDoc swAssembly = ((AssemblyDoc)(swDoc));

            



            foreach(MechanismUnit mu in mechanismUnits)
                components.Add((Component2)((AssemblyDoc)swDoc).AddComponent2(mu.FileName, mu.X, mu.Y, mu.Z));
            
            for (int i = 0; i < mechanismUnits.Count; i++)
            {
                MechanismUnit mu = mechanismUnits[i];

                swDoc.SetPickMode();
                if (i == 0)
                {
                    boolstatus = swDoc.Extension.SelectByID2("clockMechanismUnit1-1@clockMechanism", "COMPONENT", 0, 0, 0, false, 0, null, 0);
                    swAssembly.UnfixComponent();
                    swDoc.ClearSelection2(true);
                }
                
                //wi¹zania p³aszczyzna przednia-oœ
                boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna przednia", "PLANE", 0, 0, 0, true, 1, null, 0);
                boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic1@clockMechanismUnit"+mu.MUID+"-1@clockMechanism", "EXTSKETCHSEGMENT", mu.X, mu.Y+0.071264831374360668, mu.Z, true, 1, null, 0);
                Mate2 myMate = null;
                swAssembly = ((AssemblyDoc)(swDoc));
                myMate = ((Mate2)(swAssembly.AddMate5(0, -1, false, 1.1655880690436549e-005, 0.001, 0.001, 0.001, 0.001, 0.52359877559830004, 0.52359877559830004, 0.52359877559830004, false, false, 0, out longstatus)));
                swDoc.ClearSelection2(true);
                swDoc.EditRebuild3();

                //wi¹zania prostopadle oœ-p³aszczyzna górna
                boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic1@clockMechanismUnit" + mu.MUID + "-1@clockMechanism", "EXTSKETCHSEGMENT", mu.X, mu.Y + 0.071264831374360668, mu.Z, true, 1, null, 0);
                boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna górna", "PLANE", 0, 0, 0, true, 1, null, 0);
                swAssembly = ((AssemblyDoc)(swDoc));
                myMate = ((Mate2)(swAssembly.AddMate5(2, -1, false, 0, 0.001, 0.001, 0.001, 0.001, 0.52359877559830004, 0.52359877559830004, 0.52359877559830004, false, false, 0, out longstatus)));
                swDoc.ClearSelection2(true);
                swDoc.EditRebuild3();


                //wi¹zania odlag³oœci p³aszczyzn od œrodków uk³adów wspó³rzêdnych
                
                swDoc = ((ModelDoc2)(swApp.ActiveDoc));
                boolstatus = swDoc.Extension.SelectByID2("Point1@Pocz¹tek uk³adu wspó³rzêdnych@clockMechanismUnit"+mu.MUID+"-1@clockMechanism", "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
                boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna górna", "PLANE", 0, 0, 0, true, 1, null, 0);
                swAssembly = ((AssemblyDoc)(swDoc));
                myMate = ((Mate2)(swAssembly.AddMate5(5, -1, true, mu.Y, mu.Y, mu.Y, 0.001, 0.001, 0.52359877559830004, 0.52359877559830004, 0.52359877559830004, false, false, 0, out longstatus)));
                swDoc.ClearSelection2(true);
                swDoc.EditRebuild3();
                // 
                boolstatus = swDoc.Extension.SelectByID2("Point1@Pocz¹tek uk³adu wspó³rzêdnych@clockMechanismUnit" + mu.MUID + "-1@clockMechanism", "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
                boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna prawa", "PLANE", 0, 0, 0, true, 1, null, 0);
                swAssembly = ((AssemblyDoc)(swDoc));
                myMate = ((Mate2)(swAssembly.AddMate5(5, -1, false, mu.X, mu.X, mu.X, 0.001, 0.001, 0.52359877559830004, 0.52359877559830004, 0.52359877559830004, false, false, 0, out longstatus)));
                swDoc.ClearSelection2(true);
                swDoc.EditRebuild3();
               
           }

           swDoc = ((ModelDoc2)(swApp.ActiveDoc));
           boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic1@clockMechanismUnit1-1@clockMechanism", "EXTSKETCHSEGMENT", mechanismUnits[0].X, mechanismUnits[0].Y + 0.063800666147216134, mechanismUnits[0].Z, true, 1, null, 0);
           boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic1@clockMechanismUnit2-1@clockMechanism", "EXTSKETCHSEGMENT", mechanismUnits[1].X, mechanismUnits[1].Y + 0.012120404370224364, mechanismUnits[1].Z, true, 1, null, 0);
           Mate2 myMate2;
           swAssembly = ((AssemblyDoc)(swDoc));
           myMate2 = ((Mate2)(swAssembly.AddMate5(10, -1, false, 0.069999999999999993, 0.001, 0.001, (int)Math.Round(r2 / 0.05 * 30), (int)Math.Round(r1 / 0.05 * 30), 0, 0.52359877559830004, 0.52359877559830004, false, false, 0, out longstatus)));
           swDoc.ClearSelection2(true);
           swDoc.EditRebuild3();
            swDoc.Save();
            swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic1@clockMechanismUnit2-1@clockMechanism", "EXTSKETCHSEGMENT", mechanismUnits[0].X, mechanismUnits[1].Y + 0.063800666147216134, mechanismUnits[1].Z, true, 1, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic1@clockMechanismUnit3-1@clockMechanism", "EXTSKETCHSEGMENT", mechanismUnits[1].X, mechanismUnits[2].Y + 0.012120404370224364, mechanismUnits[2].Z, true, 1, null, 0);
      
            swAssembly = ((AssemblyDoc)(swDoc));
            myMate2 = ((Mate2)(swAssembly.AddMate5(10, -1, false, 0.069999999999999993, 0.001, 0.001, (int)Math.Round(r2 / 0.05 * 30), (int)Math.Round(r1 / 0.05 * 30), 0, 0.52359877559830004, 0.52359877559830004, false, false, 0, out longstatus)));
            swDoc.ClearSelection2(true);
            swDoc.EditRebuild3();
            swDoc.Save();
        }
    }
}
