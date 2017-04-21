using System;
using System.Collections.Generic;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;

namespace Macro2.csproj
{
    public class Gearwheel
    {
        private double r = 0.05;
        private double h_2 = 0.01;
    
        private string fileName;
        public string FileName
        {
            get { return fileName; }
        }
        public double R
        {
            get { return r; }
        }

         public Gearwheel(double _r, double _h, string _fileName)
        {
            r = _r;
            h_2 = _h / 2;
            fileName = _fileName;
        }


        public void makeGearwheelPart(SldWorks swApp)
        {
            ModelDoc2 swDoc = ((ModelDoc2)(swApp.NewPart()));
            ModelView myModelView = ((ModelView)(swDoc.ActiveView)); //aktywny widok
            myModelView.FrameState = ((int)(swWindowState_e.swWindowMaximized));

            //wybieramy p³aszczyznê przedni¹
            bool boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna górna", "PLANE", 0, 0, 0, false, 0, null, 0);
            //robimy na niej szkic
            swDoc.SketchManager.InsertSketch(true);
            //odznaczamy wszystko
            swDoc.ClearSelection2(true);
            //rysujemy kó³ko
            SketchSegment skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0, 0, 0, r, 0, 0)));
            //odznaczamy wszystko
            swDoc.ClearSelection2(true);
            ////nowy szkic
            swDoc.SketchManager.InsertSketch(true);

            //dodanie wyci¹gniêcia-bazy
            boolstatus = swDoc.Extension.SelectByID2("Szkic1", "SKETCH", 0, 0, 0, false, 4, null, 0);
            Feature myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(false, false, false, 0, 0, h_2, h_2, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            

            //szkic szczeliny
            boolstatus = swDoc.Extension.SelectByID2("P³aszczyzna górna", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCenterLine(0, 0, 0, 0, r * 2, 0)));
            swDoc.SetPickMode();
            swDoc.ClearSelection2(true);
            swDoc.SetPickMode();
            // 
            double[] points = new double[12];
            points[0] = 0;
            points[1] = r - 0.010;
            points[2] = 0;

            points[3] = 0.0019;
            points[4] = r - 0.0087;
            points[5] = 0;

            points[6] = 0.0035;
            points[7] = r;
            points[8] = 0;

            points[9] = 0.0059;
            points[10] = r + 0.0011;
            points[11] = 0;

            Array pointsArray = points;
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateSpline(pointsArray)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Spline1", "SKETCHSEGMENT", 0.0020095885476130518, 0.050715773715176618, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", -5.0469609751323252e-005, 0.051673940299997258, 0, true, 0, null, 0);
            swDoc.SketchMirror();
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(-0.0059, r + 0.0011, 0, 0.0059, r + 0.0011, 0)));
            swDoc.ClearSelection2(true);
            swDoc.SetPickMode();
            boolstatus = swDoc.Extension.SelectByID2("Splajn2", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(1, -0.003164511010418404, 0.051386490324551067, 0);
            swDoc.SetPickMode();
            boolstatus = swDoc.Extension.SelectByID2("Splajn1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(1, 0.0033510217663619418, 0.051434398653792096, 0);
            swDoc.SetPickMode();
            boolstatus = swDoc.Extension.SelectByID2("Linia2", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(1, -0.0035956859735876909, 0.0511, 0);
            swDoc.SetPickMode();
            boolstatus = swDoc.Extension.SelectByID2("Linia2", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(1, 0.0035426550833260739, 0.0511, 0);

            //wyciêcie szczeliny

            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Szkic2", "SKETCH", 0, 0, 0, false, 4, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(false, false, false, 0, 0, 0.01, 0.01, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            //szyk ko³owy szczelin
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Wytnij-wyci¹gniêcie1", "BODYFEATURE", 0, 0, 0, false, 4, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("", "FACE", 0, -r, 0, true, 1, null, 0);

            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCircularPattern4((int)Math.Round(r / 0.05 * 30), 6.2831853071796004, false, "NULL", false, true, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            //ODZNACZANIE WSZYSTKIEGO
            swDoc.ClearSelection2(true);
            //zapisanie do pliku
            swDoc.SaveAs(fileName);
         
           
        }

    }
}
