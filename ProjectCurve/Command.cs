#region Namespaces
using System;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace ProjectCurve
{
  [Transaction( TransactionMode.Manual )]
  public class Command : IExternalCommand
  {

    /// <summary>
    /// A selection filter for a planar face.
    /// Only planar faces are allowed to be picked.
    /// From Revit SDK Samples/Selections/CS/SelectionFilters.cs
    /// </summary>
    public class PlanarFaceFilter : ISelectionFilter
    {
      // Revit document.
      Document m_doc = null;

      /// <summary>
      /// Filter constructor, initialise the document.
      /// </summary>
      /// <param name="doc">The document.</param>
      public PlanarFaceFilter( Document doc )
      {
        m_doc = doc;
      }

      /// <summary>
      /// Allow all the element to be selected
      /// </summary>
      /// <param name="e">A candidate element in selection operation.</param>
      /// <returns>Return true to allow the user to select this candidate element.</returns>
      public bool AllowElement( Element e )
      {
        return true;
      }

      /// <summary>
      /// Allow planar face reference to be selected
      /// </summary>
      /// <param name="r">A candidate reference in selection operation.</param>
      /// <param name="p">The 3D position of the mouse on the candidate reference.</param>
      /// <returns>Return true for planar face reference. Return false for non-planar face reference.</returns>
      public bool AllowReference( Reference r, XYZ p )
      {
        GeometryObject geoObject = m_doc.GetElement( r )
          .GetGeometryObjectFromReference( r );

        return geoObject != null 
          && geoObject is PlanarFace;
      }
    }

    /// <summary>
    /// Selection filter for elements havin a location curve.
    /// </summary>
    class LocationCurvePickFilter : ISelectionFilter
    {
      public bool AllowElement( Element e )
      {
        return e.Location is LocationCurve;
      }

      public bool AllowReference( Reference r, XYZ p )
      {
        return true;
      }
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      Selection sel = uidoc.Selection;

      // Pick a planar projection surface;
      // We could also project onto a non-planar 
      // surface, but:
      // - The straight curve segments would maybe 
      //   not all lie in the curved surface
      // - We could not use model lines in a single 
      //   (planar) sketch plane to represent them

      PlanarFace face = null;

        try
        {
          Reference r = sel.PickObject(
            ObjectType.Face, new PlanarFaceFilter( doc ),
            "Please pick a projection surface" );

          GeometryObject geoObject = doc.GetElement( r )
            .GetGeometryObjectFromReference( r );

          face = geoObject as PlanarFace;
        }
        catch( Autodesk.Revit.Exceptions.OperationCanceledException )
        {
          message = "Selection cancelled.";
          return Result.Cancelled;
        }

      // Pick the curves to project

      List<Curve> curves = new List<Curve>();

      try
      {
        IList<Reference> refs = sel.PickObjects(
          ObjectType.Element, new LocationCurvePickFilter(),
          "Please pick curves to project" );

        foreach( Reference r in refs )
        {
          Element e = doc.GetElement( r.ElementId );
          LocationCurve lc = e.Location as LocationCurve;
          curves.Add( lc.Curve );
        }
      }
      catch( Autodesk.Revit.Exceptions.OperationCanceledException )
      {
        message = "Selection cancelled.";
        return Result.Cancelled;
      }

      using( Transaction tx = new Transaction( doc ) )
      {
        tx.Start( "Transaction Name" );

        Plane plane = Plane.CreateByNormalAndOrigin(
          face.FaceNormal, face.Origin );

        SketchPlane sketch = SketchPlane.Create(
          doc, plane );

        foreach( Curve c in curves )
        {
          IList<XYZ> pts = c.Tessellate();
          int n = pts.Count;
          XYZ p = pts[0];
          for( int i = 1; i < n; ++i )
          {
            XYZ q = pts[i];
            Line line = Line.CreateBound( p, q );
            doc.Create.NewModelCurve( line, sketch );
            p = q;
          }
        }
        tx.Commit();
      }
      return Result.Succeeded;
    }
  }
}
