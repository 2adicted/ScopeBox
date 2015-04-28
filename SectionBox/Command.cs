#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Excel = Microsoft.Office.Interop.Excel; 
using System.Windows.Forms;
#endregion

namespace SectionBox
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        const double _eps = 1.0e-9;
        static private string name_dummy;
        private FamilySymbol m_TitleBlocks;
        private double TITLEBAR = 0.2;
        private double GOLDENSECTION = 0.618;
        private double m_rows;

        /// <summary>
        /// Return true if the given real number is almost zero.
        /// </summary>
        static bool IsAlmostSero(double a)
        {
            return _eps > Math.Abs(a);
        }

        /// <summary>
        /// Return true if the given vector is almost vertical.
        /// </summary>
        static bool IsVertical(XYZ v)
        {
            return IsAlmostSero(v.X) && IsAlmostSero(v.Y);
        }

        /// <summary>
        /// Return true if v and w are non-zero and perpendicular.
        /// </summary>
        static bool IsPerpendicular(XYZ v, XYZ w)
        {
            double a = v.GetLength();
            double b = w.GetLength();
            double c = Math.Abs(v.DotProduct(w));
            return _eps < a
              && _eps < b
              && _eps > c;

            // To take the relative lengths of a and b into 
            // account, you can scale the whole test, e.g
            // c * c < _eps * a * b... can you?
        }

        /// <summary>
        /// Return the signed volume of the paralleliped 
        /// spanned by the vectors a, b and c. In German, 
        /// this is also known as Spatprodukt.
        /// </summary>
        static double SignedParallelipedVolume(
          XYZ a,
          XYZ b,
          XYZ c)
        {
            return a.CrossProduct(b).DotProduct(c);
        }

        /// <summary>
        /// Return true if the three vectors a, b and c 
        /// form a right handed coordinate system, i.e.
        /// the signed volume of the paralleliped spanned 
        /// by them is positive.
        /// </summary>
        bool IsRightHanded(XYZ a, XYZ b, XYZ c)
        {
            return 0 < SignedParallelipedVolume(a, b, c);
        }


        public void loadExcelSheet(String filename)
        {
            Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();

            object misValue = System.Reflection.Missing.Value;
            Excel.Workbook workbook = xlapp.Workbooks.Open(filename);
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = worksheet.UsedRange;;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            str = (string)(range.Cells[5, 2] as Excel.Range).Value2;
            set_str(str);

            rCnt = range.Rows.Count;
            cCnt = range.Columns.Count;


            releaseObject(worksheet);
            releaseObject(workbook);
            releaseObject(xlapp);
        }

        private void set_str(string str)
        {
            name_dummy = str;
        }

        public string get_str()
        {
            return name_dummy;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        } 

        BoundingBoxXYZ GenerateBoundingBoxXYZ(Element scopeBox)
        {
            BoundingBoxXYZ sbox = scopeBox.get_BoundingBox(null);
            BoundingBoxXYZ mbox = new BoundingBoxXYZ();
            mbox.Min = sbox.Min;
            mbox.Max = sbox.Max;
            mbox.Enabled = true;
            Transform transform = TransformFromScopeBox(scopeBox);
            mbox.Transform = transform;
            return mbox;
        }

        /// <summary>
        /// Retrieve all available title blocks in the currently active document.
        /// </summary>
        /// <param name="doc">the currently active document</param>
        private void GetTitleBlocks(Document doc)
        {
            m_TitleBlocks = new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_TitleBlocks)
            .FirstElement() as FamilySymbol;

            if (null == m_TitleBlocks)
            {
                throw new InvalidOperationException("There is no title block to generate sheet.");
            }           
        }

        /// <summary>
        /// Generate sheet in active document.
        /// </summary>
        /// <param name="doc">the currently active document</param>
        public void GenerateSheet(Document doc, Autodesk.Revit.DB.View m_selectedView)
        {
            if (null == doc)
            {
                throw new ArgumentNullException("doc");
            }

            ViewSheet sheet = ViewSheet.Create(doc, m_TitleBlocks.Id);
            ViewSet m_selectedViews = new ViewSet();

            sheet.Name = get_str();
            
            m_selectedViews.Insert(m_selectedView);
            PlaceViews(m_selectedViews, sheet);
        }

        /// <summary>
        /// Place all selected views on this sheet's appropriate location.
        /// </summary>
        /// <param name="views">all selected views</param>
        /// <param name="sheet">all views located sheet</param>
        private void PlaceViews(ViewSet views, ViewSheet sheet)
        {
            double xDistance = 0;
            double yDistance = 0;
            CalculateDistance(sheet.Outline, views.Size, ref xDistance, ref yDistance);
            Autodesk.Revit.DB.UV origin = GetOffSet(sheet.Outline, xDistance, yDistance);
            //Autodesk.Revit.DB.UV temp = new Autodesk.Revit.DB.UV (origin.U, origin.V);
            double tempU = origin.U;
            double tempV = origin.V;
            int n = 1;
            foreach (Autodesk.Revit.DB.View v in views)
            {
                Autodesk.Revit.DB.UV location = new Autodesk.Revit.DB.UV(tempU, tempV);
                Autodesk.Revit.DB.View view = v;
                Rescale(view, xDistance, yDistance);
                try
                {
                    //sheet.AddView(view, location);
                    Viewport.Create(view.Document, sheet.Id, view.Id, new XYZ(location.U, location.V, 0));
                }
                catch (ArgumentException /*ae*/)
                {
                    throw new InvalidOperationException("The view '" + view.Name +
                        "' can't be added, it may have already been placed in another sheet.");
                }

                if (0 != n++ % m_rows)
                {
                    tempU = tempU + xDistance * (1 - TITLEBAR);
                }
                else
                {
                    tempU = origin.U;
                    tempV = tempV + yDistance;
                }
            }
        }

        

        /// <summary>
        /// Retrieve the appropriate origin.
        /// </summary>
        /// <param name="bBox"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        private Autodesk.Revit.DB.UV GetOffSet(BoundingBoxUV bBox, double x, double y)
        {
            return new Autodesk.Revit.DB.UV(bBox.Min.U + x * GOLDENSECTION, bBox.Min.V + y * GOLDENSECTION);
        }

        /// <summary>
        /// Calculate the appropriate distance between the views lay on the sheet.
        /// </summary>
        /// <param name="bBox">The outline of sheet.</param>
        /// <param name="amount">Amount of views.</param>
        /// <param name="x">Distance in x axis between each view</param>
        /// <param name="y">Distance in y axis between each view</param>
        private void CalculateDistance(BoundingBoxUV bBox, int amount, ref double x, ref double y)
        {
            double xLength = (bBox.Max.U - bBox.Min.U) * (1 - TITLEBAR);
            double yLength = (bBox.Max.V - bBox.Min.V);

            //calculate appropriate rows numbers.
            double result = Math.Sqrt(amount);

            while (0 < (result - (int)result))
            {
                amount = amount + 1;
                result = Math.Sqrt(amount);
            }
            m_rows = result;
            double area = xLength * yLength / amount;

            //calculate appropriate distance between the views.
            if (bBox.Max.U > bBox.Max.V)
            {
                x = Math.Sqrt(area / GOLDENSECTION);
                y = GOLDENSECTION * x;
            }
            else
            {
                y = Math.Sqrt(area / GOLDENSECTION);
                x = GOLDENSECTION * y;
            }
        }

        /// <summary>
        /// Rescale the view's Scale value for suitable.
        /// </summary>
        /// <param name="view">The view to be located on sheet.</param>
        /// <param name="x">Distance in x axis between each view</param>
        /// <param name="y">Distance in y axis between each view</param>
        static private void Rescale(Autodesk.Revit.DB.View view, double x, double y)
        {
            double Rescale = 2;
            Autodesk.Revit.DB.UV outline = new Autodesk.Revit.DB.UV(view.Outline.Max.U - view.Outline.Min.U,
                view.Outline.Max.V - view.Outline.Min.V);

            if (outline.U > outline.V)
            {
                Rescale = outline.U / x * Rescale;
            }
            else
            {
                Rescale = outline.V / y * Rescale;
            }

            if (1 != view.Scale && 0 != Rescale)
            {
                view.Scale = (int)(view.Scale * Rescale);
            }
        }

        /// <summary>
        /// Generating the transform for one side of the box
        /// </summary>
        Transform TransformFromScopeBox(
          Element scopeBox)
        {
            Document doc = scopeBox.Document;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;
           
            // Retrieve scope box geometry, 
            // consisting of exactly twelve lines.

            Options opt = app.Create.NewGeometryOptions();
            GeometryElement geo = scopeBox.get_Geometry(opt);
            int n = geo.Count<GeometryObject>();

            if (12 != n)
            {
                throw new ArgumentException("Expected exactly"
                  + " 12 lines in scope box geometry");
            }

            // Determine origin as the bottom endpoint of 
            // the edge closest to the viewer, and vz as the 
            // vertical upwards pointing vector emanating
            // from it. (Todo: if several edges are equally 
            // close, pick the leftmost one, assuming the 
            // given view direction and Z is upwards.)
            Line line = null;
            Transform transform  = null;
            foreach (GeometryObject obj in geo)
            {
                Debug.Assert(obj is Line,
                  "expected only lines in scope box geometry");

                line = obj as Line;

                XYZ p, q, v;
                
                p = line.GetEndPoint(0);
                q = line.GetEndPoint(1);
                v = q - p;

                if (!IsVertical(v))
                {
                    break;
                }
            }

            transform = Transform.Identity;
            XYZ mPoint = XYZMath.FindMidPoint(line.GetEndPoint(0), line.GetEndPoint(1));

            transform.Origin = mPoint;

            // At last find out the directions of the created view, and set it as Basis property.
            Autodesk.Revit.DB.XYZ basisZ = XYZMath.FindDirection(line.GetEndPoint(0), line.GetEndPoint(1));
            Autodesk.Revit.DB.XYZ basisX = XYZMath.FindRightDirection(basisZ);
            Autodesk.Revit.DB.XYZ basisY = XYZMath.FindUpDirection(basisZ);


            transform.set_Basis(0, basisX);
            transform.set_Basis(1, basisY);
            transform.set_Basis(2, basisZ);
            return transform;         
        }


        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Selection sel = uidoc.Selection;

            //performCheck(doc, message);

            var form_to_show = new LoadExcel(uiapp);

            form_to_show.ShowDialog();


            View3D view = Get3dView(doc);
            
            if (null == view)
            {
                message = "Please run this command in a 3D view.";
                return Result.Failed;
            }

            uidoc.ActiveView = view;

            Element scopeBox
              = new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                .WhereElementIsNotElementType()
                .FirstElement();
            if (null == scopeBox)
            {
                message = "Please make sure you have at least one ScopeBox in the project.";
                return Result.Failed;
            }
            ElementId scopeBoxName = scopeBox.Id;
            
            BoundingBoxXYZ viewScopeBox
              = GenerateBoundingBoxXYZ(scopeBox);
           
            //set up view family type - have to play with this one more
            ViewFamilyType vft
      = new FilteredElementCollector(doc)
        .OfClass(typeof(ViewFamilyType))
        .Cast<ViewFamilyType>()
        .FirstOrDefault<ViewFamilyType>(x =>
          ViewFamily.Section == x.ViewFamily);
            string filenametemp = get_str();
            //make the transaction; should add more sections here later;
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Move And Resize Section Box");
                ViewSection section = ViewSection.CreateSection(doc, vft.Id, viewScopeBox);
                section.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(scopeBoxName);
                section.get_Parameter(BuiltInParameter.VIEW_NAME).Set(get_str());
                GetTitleBlocks(doc);
                GenerateSheet(doc, section);
                //view.SetSectionBox(viewScopeBox);
                tx.Commit();
            }

            return Result.Succeeded;
        }


        private void performCheck(Document doc, string message)
        {
            View3D view = doc.ActiveView as View3D;
            if(null == view)
            {
                message = "Please run this command in a 3D view.";
            }
        }

        /// <summary>
        /// Retrieve a suitable 3D view from document.
        /// </summary>
        View3D Get3dView(Document doc)
        {
            FilteredElementCollector collector
              = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D));

            foreach (View3D v in collector)
            {
                Debug.Assert(null != v,
                  "never expected a null view to be returned"
                  + " from filtered element collector");

                // Skip view template here because view 
                // templates are invisible in project 
                // browser

                if (!v.IsTemplate)
                {
                    return v;
                }
            }
            return null;
        }

    }
}
