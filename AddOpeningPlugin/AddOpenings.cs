using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOpeningPlugin
{
    [Transaction(TransactionMode.Manual)]
    public class AddOpenings : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document mepDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            double offset = UnitUtils.ConvertToInternalUnits(30, UnitTypeId.Millimeters);

            if (mepDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл");
                return Result.Cancelled;
            }

            FamilySymbol openingType = new FilteredElementCollector(arDoc)
                                        .OfClass(typeof(FamilySymbol))
                                        .OfCategory(BuiltInCategory.OST_GenericModel)
                                        .OfType<FamilySymbol>()
                                        .Where(x => x.FamilyName.Equals("Opening"))
                                        .FirstOrDefault();

            if (openingType == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство отверстия");
                return Result.Cancelled;
            }


            List<Duct> ducts = new FilteredElementCollector(mepDoc)
                                    .OfClass(typeof(Duct))
                                    .OfType<Duct>()
                                    .ToList();
            List<Pipe> pipes = new FilteredElementCollector(mepDoc)
                                    .OfClass(typeof(Pipe))
                                    .OfType<Pipe>()
                                    .ToList();
            View view3D = new FilteredElementCollector(arDoc)
                                .OfClass(typeof(View3D))
                                .OfType<View3D>()
                                .Where(x => !x.IsTemplate)
                                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector intersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D as View3D);


            Transaction ts_duct = new Transaction(arDoc, "Openings for ducts");
            ts_duct.Start();

            if (!openingType.IsActive)
                openingType.Activate();
            foreach (Duct duct in ducts)
            {
                Line d_line = (duct.Location as LocationCurve).Curve as Line;
                XYZ d_point = d_line.GetEndPoint(0);
                XYZ d_direction = d_line.Direction;
                List<ReferenceWithContext> d_intersections = intersector.Find(d_point, d_direction)
                                                           .Where(x => x.Proximity <= d_line.Length)
                                                           .Distinct(new ReferenceWithContextElemetEqualityComparer())
                                                           .ToList();
                foreach (ReferenceWithContext d_refer in d_intersections)
                {
                    double d_proximity = d_refer.Proximity;
                    Reference d_reference = d_refer.GetReference();
                    Wall d_wall = arDoc.GetElement(d_reference.ElementId) as Wall;
                    Level d_level = arDoc.GetElement(d_wall.LevelId) as Level;
                    XYZ d_pointOpening = d_point + (d_direction * d_proximity);

                    FamilyInstance d_opening = arDoc.Create.NewFamilyInstance(d_pointOpening, openingType, d_wall, d_level, StructuralType.NonStructural);
                    Parameter d_widthOpening = d_opening.LookupParameter("Width");
                    Parameter d_heightOpening = d_opening.LookupParameter("Height");

                    d_widthOpening.Set(duct.Diameter + offset);
                    d_heightOpening.Set(duct.Diameter + offset);
                }
            }
            ts_duct.Commit();


            Transaction ts_pipe = new Transaction(arDoc, "Openings for pipes");
            ts_pipe.Start();

            if (!openingType.IsActive)
                openingType.Activate();
            foreach (Pipe pipe in pipes)
            {
                Line p_line = (pipe.Location as LocationCurve).Curve as Line;
                XYZ p_point = p_line.GetEndPoint(0);
                XYZ p_direction = p_line.Direction;
                List<ReferenceWithContext> p_intersections = intersector.Find(p_point, p_direction)
                                                           .Where(x => x.Proximity <= p_line.Length)
                                                           .Distinct(new ReferenceWithContextElemetEqualityComparer())
                                                           .ToList();
                foreach (ReferenceWithContext p_refer in p_intersections)
                {
                    double p_proximity = p_refer.Proximity;
                    Reference p_reference = p_refer.GetReference();
                    Wall p_wall = arDoc.GetElement(p_reference.ElementId) as Wall;
                    Level p_level = arDoc.GetElement(p_wall.LevelId) as Level;
                    XYZ p_pointOpening = p_point + (p_direction * p_proximity);

                    FamilyInstance p_opening = arDoc.Create.NewFamilyInstance(p_pointOpening, openingType, p_wall, p_level, StructuralType.NonStructural);
                    Parameter p_widthOpening = p_opening.LookupParameter("Width");
                    Parameter p_heightOpening = p_opening.LookupParameter("Height");

                    p_widthOpening.Set(pipe.Diameter + offset);
                    p_heightOpening.Set(pipe.Diameter + offset);
                }
            }
            ts_pipe.Commit();
            return Result.Succeeded;
        }
    }

    public class ReferenceWithContextElemetEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y))
                return true;
            if (ReferenceEquals(null, x))
                return false;
            if (ReferenceEquals(null, y))
                return false;

            var xReference = x.GetReference();
            var yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId &&
                   xReference.ElementId == yReference.ElementId;
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}
