using System;
using System.Collections.Generic;
using System.Text;
using SolidWorks.Interop.sldworks;
using System.Windows.Forms;

namespace Macro2.csproj
{
    public class MechanismUnit
    {
        private Gearwheel g1, g2;
        private Stick stick;
        private double x, y, z;
        private double h1, h2;
        private double r1, r2, stickR, wheelWidth;
        private string folderPath;
        private string fileName;
        private int muID;
        private ModelDoc2 swDoc;
        private string assemblyName = "clockMechanismUnit";
        private bool g2Present = true;

        public Gearwheel G1 { get { return g1; } }
        public Gearwheel G2 { get { return g2; } }
        public Stick Stick { get { return stick; } }
        public double X { get { return x; } }
        public double Y { get { return y; } }
        public double Z { get { return z; } }
        public double H1 { get { return h1; } }
        public double H2 { get { return h2; } }
        public int MUID { get { return muID; } }
        public string FileName { get{ return fileName;}  }

        public MechanismUnit(double _r1, double _r2, bool _g2Present, double _wheelWidth, double _stickR, double _x, double _y, double _z, double _h1, double _h2, string _folderPath, int _mechanismUnitID, SldWorks swApp)
        {

            g2Present = _g2Present;
            r1 = _r1;
            r2 = _r2;
            wheelWidth = _wheelWidth;
            stickR = _stickR;
            x = _x;
            y = _y;
            z = _z;
            h1 = _h1;
            h2 = _h2;
            folderPath = _folderPath;
            muID = _mechanismUnitID;
            swDoc = ((ModelDoc2)(swApp.NewAssembly()));
            assemblyName = assemblyName + muID;
            fileName = folderPath + "\\" + assemblyName + ".SLDASM";
            //wstêpny zapis
            swDoc.SaveAs(FileName);
            createMechanismUnit(swApp);
        }

        private void createMechanismUnit(SldWorks swApp)
        {
            g1 = new Gearwheel(r1, wheelWidth, folderPath + "\\gearwheel1.SLDPRT");
            g2 = new Gearwheel(r2, wheelWidth, folderPath + "\\gearwheel2.SLDPRT");

            stick = new Stick(stickR, h1, folderPath + "\\stick.SLDPRT");

            if (muID == 1)
            {
                g1.makeGearwheelPart(swApp);
                g2.makeGearwheelPart(swApp);
                stick.makeStickPart(swApp);
            }
        }

        public void makeMechanismUnit(SldWorks swApp)
        {
            bool boolstatus;
            AssemblyDoc swAssembly = ((AssemblyDoc)(swDoc));
            int errors = 0;

            swApp.ActivateDoc3(folderPath + "\\" + assemblyName + ".SLDASM", true, (int)SolidWorks.Interop.swconst.swRebuildOnActivation_e.swUserDecision, ref errors);
            swDoc.ClearSelection2(true);

            
            ((AssemblyDoc)swDoc).AddComponent2(G1.FileName, 0, 0, 0);
            if(g2Present)
                ((AssemblyDoc)swDoc).AddComponent2(G2.FileName, 0,0,0);
            ((AssemblyDoc)swDoc).AddComponent2(Stick.FileName, 0,0,0);
                

            swDoc.SetPickMode();

            boolstatus = swDoc.Extension.SelectByID2("gearwheel1-1@" + assemblyName, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swAssembly.UnfixComponent();
            swDoc.ClearSelection2(true);


            //ustawianie sticka
            SketchSegment skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCenterLine(0, 0, 0, 0, 0.1, 0)));
            swDoc.SetPickMode();
            swDoc.ClearSelection2(true);
            SketchPoint skPoint = ((SketchPoint)(swDoc.SketchManager.CreatePoint(0, 0, 0)));
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Point1@Szkic1", "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Point1@Pocz¹tek uk³adu wspó³rzêdnych@stick-" + 1 + "@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);

            Mate2 myMate = null;
            swAssembly = ((AssemblyDoc)(swDoc));
            int longstatus;
            myMate = ((Mate2)(swAssembly.AddMate5(5, -1, false, 0, 0, 0, 0, 0, 0, 0, 0, false, false, 0, out longstatus)));

            swDoc.ClearSelection2(true);
            swDoc.EditRebuild3();

            boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic" + 1, "EXTSKETCHSEGMENT", 0, 0.01, 0, true, 1, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna przednia@stick-" + 1 + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
            swAssembly = ((AssemblyDoc)(swDoc));
            myMate = ((Mate2)(swAssembly.AddMate5(0, -1, false, 0, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
            swDoc.ClearSelection2(true);
            swDoc.EditRebuild3();
            // 
            boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic" + 1, "EXTSKETCHSEGMENT", 0, 0.01, 0, true, 1, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna prawa@stick-" + 1 + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
            swAssembly = ((AssemblyDoc)(swDoc));
            myMate = ((Mate2)(swAssembly.AddMate5(0, -1, false, 0, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
            swDoc.ClearSelection2(true);
            swDoc.EditRebuild3();
            // 

            //wi¹zania dla 1. zêbatki
            boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna górna@gearwheel1-" + 1 + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic" + 1, "EXTSKETCHSEGMENT", 0, 0.032539520735980425, 0, true, 1, null, 0);
            myMate = ((Mate2)(swAssembly.AddMate5(2, -1, false, 0.0083743333390084666, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
            swDoc.ClearSelection2(true);
            swDoc.EditRebuild3();
            // 
            boolstatus = swDoc.Extension.SelectByID2("Point1@Pocz¹tek uk³adu wspó³rzêdnych@gearwheel1-" + 1 + "@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic" + 1, "EXTSKETCHSEGMENT", 0, 0.01, 0, true, 1, null, 0);
            myMate = ((Mate2)(swAssembly.AddMate5(0, -1, false, 0.020760541438047318, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
            swDoc.ClearSelection2(true);
            swDoc.EditRebuild3();
            // 
            boolstatus = swDoc.Extension.SelectByID2("Point1@Pocz¹tek uk³adu wspó³rzêdnych@gearwheel1-" + 1 + "@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Point1@Szkic" + 1, "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
            swAssembly = ((AssemblyDoc)(swDoc));
            myMate = ((Mate2)(swAssembly.AddMate5(5, -1, false, H1, H1, H1, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
            swDoc.ClearSelection2(true);
            swDoc.EditRebuild3();
            //
            boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna przednia@stick-" + 1 + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna przednia@gearwheel1-" + 1 + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
            myMate = ((Mate2)(swAssembly.AddMate5(0, 0, false, 0, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
            swDoc.ClearSelection2(true);
            swDoc.EditRebuild3();

            if (g2Present)
            {
                //wi¹zania dla 2. zêbatki
                boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna górna@gearwheel2-" + 1 + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
                boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic" + 1, "EXTSKETCHSEGMENT", 0, 0.01, 0, true, 1, null, 0);
                myMate = ((Mate2)(swAssembly.AddMate5(2, -1, false, 0.0083743333390084666, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
                swDoc.ClearSelection2(true);
                swDoc.EditRebuild3();
                // 
                boolstatus = swDoc.Extension.SelectByID2("Point1@Pocz¹tek uk³adu wspó³rzêdnych@gearwheel2-" + 1 + "@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
                boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic" + 1, "EXTSKETCHSEGMENT", 0, 0.01, 0, true, 1, null, 0);
                myMate = ((Mate2)(swAssembly.AddMate5(0, -1, false, 0.020760541438047318, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
                swDoc.ClearSelection2(true);
                swDoc.EditRebuild3();
                // 
                boolstatus = swDoc.Extension.SelectByID2("Point1@Pocz¹tek uk³adu wspó³rzêdnych@gearwheel2-" + 1 + "@" + assemblyName, "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
                boolstatus = swDoc.Extension.SelectByID2("Point1@Szkic" + 1, "EXTSKETCHPOINT", 0, 0, 0, true, 1, null, 0);
                swAssembly = ((AssemblyDoc)(swDoc));
                myMate = ((Mate2)(swAssembly.AddMate5(5, -1, false, H2, H2, H2, 0.001, 0.001, 0.5, 0.5, 0.5, false, false, 0, out longstatus)));
                swDoc.ClearSelection2(true);
                swDoc.EditRebuild3();
                //
                boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna przednia@stick-" + 1 + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
                boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna przednia@gearwheel2-" + 1 + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
                myMate = ((Mate2)(swAssembly.AddMate5(0, 0, false, 0, 0.001, 0.001, 0.001, 0.001, 0, 0, 0, false, false, 0, out longstatus)));
                swDoc.ClearSelection2(true);
                swDoc.EditRebuild3();
            }
          
            //ukrycie centerLine
            boolstatus = swDoc.Extension.SelectByID2("Line1@Szkic" + 1, "EXTSKETCHSEGMENT", 0, 0.01, 0, true, 1, null, 0);
            swDoc.BlankSketch();
            swDoc.ClearSelection2(true);
            swDoc.Save();

        }
    }
}
