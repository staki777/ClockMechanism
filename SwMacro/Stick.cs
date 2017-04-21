using System;
using System.Collections.Generic;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Macro2.csproj
{
    public class Stick
    {
        private double h;
        private double r;

        private string fileName;

        public double H 
        {
            get { return h; }
        }
        public double R
        {
            get { return r; }
        }
        public string FileName
        {
            get { return fileName; }
        }

        public Stick(double _r, double _h, string _filename)
        {
            r = _r;
            h = _h;
            fileName = _filename;
        }

        public void makeStickPart(SldWorks swApp)
        {
            ModelDoc2 swDoc = ((ModelDoc2)(swApp.NewPart()));
            ModelView myModelView = ((ModelView)(swDoc.ActiveView)); //aktywny widok
            myModelView.FrameState = ((int)(swWindowState_e.swWindowMaximized));

            //wybieramy p³aszczyznê górn¹
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
            swDoc.ISelectionManager.EnableContourSelection = true;
            boolstatus = swDoc.Extension.SelectByID2("Szkic1", "SKETCHCONTOUR", 0, 0, 0, true, 4, null, 0);
            Feature myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, h, 0, false, false, false, false, 0.017453292519943334, 0.017453292519943334, false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
            //ODZNACZANIE WSZYSTKIEGO
            swDoc.ClearSelection2(true);
            //zapisanie do pliku
            swDoc.SaveAs(fileName);

        }
    }
}
