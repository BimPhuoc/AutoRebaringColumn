#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Text;
#endregion

namespace AutoRebaringColumn
{
    public static class CheckGeometry
    {
        public static Plane GetPlane(PlanarFace f)
        {
            return Plane.CreateByOriginAndBasis(f.Origin, f.XVector, f.YVector);
        }
        public static Plane GetPlaneWithBasisX(PlanarFace f, XYZ vecX)
        {
            if (!GeomUtil.IsEqual(GeomUtil.DotMatrix(vecX, f.FaceNormal), 0)) throw new Exception("VecX is not perpendicular with Normal!");
            return Plane.CreateByOriginAndBasis(f.Origin, vecX, GeomUtil.UnitVector(GeomUtil.CrossMatrix(vecX, f.FaceNormal)));
        }
        public static Plane GetPlaneWithBasisY(PlanarFace f, XYZ vecY)
        {
            if (!GeomUtil.IsEqual(GeomUtil.DotMatrix(vecY, f.FaceNormal), 0)) throw new Exception("VecY is not perpendicular with Normal!");
            return Plane.CreateByOriginAndBasis(f.Origin, GeomUtil.UnitVector(GeomUtil.CrossMatrix(vecY, f.FaceNormal)), vecY);
        }
        public static List<Curve> GetCurves(PlanarFace f)
        {
            List<Curve> curves = new List<Curve>();
            IList<CurveLoop> curveLoops = f.GetEdgesAsCurveLoops();
            foreach (CurveLoop cl in curveLoops)
            {
                foreach (Curve c in cl)
                {
                    curves.Add(c);
                }
                break;
            }
            return curves;
        }
        public static double GetSignedDistance(Plane plane, XYZ point)
        {
            XYZ v = point - plane.Origin;
            return Math.Abs(GeomUtil.DotMatrix(plane.Normal, v));
        }
        public static double GetSignedDistance(Line line, XYZ point)
        {
            if (IsPointInLineOrExtend(line, point)) return 0;
            return GeomUtil.GetLength(point, GetProjectPoint(line, point));
        }
        public static double GetSignedDistance(Curve line, XYZ point)
        {
            if (IsPointInLineOrExtend(ConvertLine(line), point)) return 0;
            return GeomUtil.GetLength(point, GetProjectPoint(line, point));
        }
        public static XYZ GetProjectPoint(Line line, XYZ point)
        {
            if (IsPointInLineOrExtend(line, point)) return point;
            XYZ vecL = GeomUtil.SubXYZ(line.GetEndPoint(1), line.GetEndPoint(0));
            XYZ vecP = GeomUtil.SubXYZ(point, line.GetEndPoint(0));
            Plane p = Plane.CreateByOriginAndBasis(line.GetEndPoint(0), GeomUtil.UnitVector(vecL), GeomUtil.UnitVector(GeomUtil.CrossMatrix(vecL, vecP)));
            return GetProjectPoint(p, point);
        }
        public static XYZ GetProjectPoint(Curve line, XYZ point)
        {
            if (IsPointInLineOrExtend(CheckGeometry.ConvertLine(line), point)) return point;
            XYZ vecL = GeomUtil.SubXYZ(line.GetEndPoint(1), line.GetEndPoint(0));
            XYZ vecP = GeomUtil.SubXYZ(point, line.GetEndPoint(0));
            Plane p = Plane.CreateByOriginAndBasis(line.GetEndPoint(0), GeomUtil.UnitVector(vecL), GeomUtil.UnitVector(GeomUtil.CrossMatrix(vecL, vecP)));
            return GetProjectPoint(p, point);
        }
        public static XYZ GetProjectPoint(Plane plane, XYZ point)
        {
            double d = GetSignedDistance(plane, point);
            XYZ q = GeomUtil.AddXYZ(point, GeomUtil.MultiplyVector(plane.Normal, d));
            return IsPointInPlane(plane, q) ? q : GeomUtil.AddXYZ(point, GeomUtil.MultiplyVector(plane.Normal, -d));
        }
        public static XYZ GetProjectPoint(PlanarFace f, XYZ point)
        {
            Plane p = GetPlane(f);
            return GetProjectPoint(p, point);
        }
        public static Curve GetProjectLine(Plane plane, Curve c)
        {
            return Line.CreateBound(GetProjectPoint(plane, c.GetEndPoint(0)), GetProjectPoint(plane, c.GetEndPoint(1)));
        }
        public static Polygon GetProjectPolygon(Plane plane, Polygon polygon)
        {
            List<Curve> cs = new List<Curve>();
            foreach (Curve c in polygon.ListCurve)
            {
                cs.Add(GetProjectLine(plane, c));
            }
            return new Polygon(cs);
        }
        public static UV Evaluate(Plane plane, XYZ point)
        {
            if (!IsPointInPlane(plane, point)) point = GetProjectPoint(plane, point);
            Plane planeOx = Plane.CreateByOriginAndBasis(plane.Origin, plane.XVec, plane.Normal);
            Plane planeOy = Plane.CreateByOriginAndBasis(plane.Origin, plane.YVec, plane.Normal);
            double lenX = GetSignedDistance(planeOy, point);
            double lenY = GetSignedDistance(planeOx, point);
            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    double tLenX = lenX * Math.Pow(-1, i + 1);
                    double tLenY = lenY * Math.Pow(-1, j + 1);
                    XYZ tPoint = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(plane.Origin, plane.XVec, tLenX), plane.YVec, tLenY);
                    if (GeomUtil.IsEqual(tPoint, point)) return new UV(tLenX, tLenY);
                }
            }
            throw new Exception("Code complier should never be here!");
        }
        public static XYZ Evaluate(Plane p, UV point)
        {
            XYZ pnt = p.Origin;
            pnt = GeomUtil.OffsetPoint(pnt, p.XVec, point.U);
            pnt = GeomUtil.OffsetPoint(pnt, p.YVec, point.V);
            return pnt;
        }
        public static UV Evaluate(PlanarFace f, XYZ point)
        {
            return Evaluate(GetPlane(f), point);
        }
        public static XYZ Evaluate(PlanarFace f, UV point)
        {
            return f.Evaluate(point);
        }
        public static UV Evaluate(Polygon f, XYZ point)
        {
            return Evaluate(GetPlane(f.Face), point);
        }
        public static XYZ Evaluate(Polygon f, UV point)
        {
            return f.Face.Evaluate(point);
        }
        public static bool IsPointInPlane(Plane plane, XYZ point)
        {
            return GeomUtil.IsEqual(GetSignedDistance(plane, point), 0) ? true : false;
        }
        public static bool IsPointInPolygon(UV p, List<UV> polygon)
        {
            double minX = polygon[0].U;
            double maxX = polygon[0].U;
            double minY = polygon[0].V;
            double maxY = polygon[0].V;
            for (int i = 1; i < polygon.Count; i++)
            {
                UV q = polygon[i];
                minX = Math.Min(q.U, minX);
                maxX = Math.Max(q.U, maxX);
                minY = Math.Min(q.V, minY);
                maxY = Math.Max(q.V, maxY);
            }

            if (p.U < minX || p.U > maxX || p.V < minY || p.V > maxY)
            {
                return false;
            }
            bool inside = false;
            for (int i = 0, j = polygon.Count - 1; i < polygon.Count; j = i++)
            {
                if ((polygon[i].V > p.V) != (polygon[j].V > p.V) &&
                     p.U < (polygon[j].U - polygon[i].U) * (p.V - polygon[i].V) / (polygon[j].V - polygon[i].V) + polygon[i].U)
                {
                    inside = !inside;
                }
            }
            return inside;
        }
        public static bool IsPointInLine(Line line, XYZ point)
        {
            if (GeomUtil.IsEqual(point, line.GetEndPoint(0)) || GeomUtil.IsEqual(point, line.GetEndPoint(1))) return true;
            if (GeomUtil.IsOppositeDirection(GeomUtil.SubXYZ(point, line.GetEndPoint(0)), GeomUtil.SubXYZ(point, line.GetEndPoint(1)))) return true;
            return false;
        }
        public static bool IsPointInLineExtend(Line line, XYZ point)
        {
            if (GeomUtil.IsEqual(point, line.GetEndPoint(0)) || GeomUtil.IsEqual(point, line.GetEndPoint(1))) return true;
            if (GeomUtil.IsSameDirection(GeomUtil.SubXYZ(point, line.GetEndPoint(0)), GeomUtil.SubXYZ(point, line.GetEndPoint(1)))) return true;
            return false;
        }
        public static bool IsPointInLineOrExtend(Line line, XYZ point)
        {
            if (GeomUtil.IsEqual(point, line.GetEndPoint(0)) || GeomUtil.IsEqual(point, line.GetEndPoint(1))) return true;
            if (GeomUtil.IsSameOrOppositeDirection(GeomUtil.SubXYZ(point, line.GetEndPoint(0)), GeomUtil.SubXYZ(point, line.GetEndPoint(1)))) return true;
            return false;
        }
        public static PointComparePolygonResult PointComparePolygon(UV p, List<UV> polygon)
        {
            bool check1 = IsPointInPolygon(p, polygon);
            for (int i = 0; i < polygon.Count; i++)
            {
                if (GeomUtil.IsEqual(p, polygon[i])) return PointComparePolygonResult.Node;

                UV vec1 = GeomUtil.SubXYZ(p, polygon[i]);
                UV vec2 = null;
                if (i != polygon.Count - 1)
                {
                    if (GeomUtil.IsEqual(p, polygon[i + 1])) continue;
                    vec2 = GeomUtil.SubXYZ(p, polygon[i + 1]);
                }
                else
                {
                    if (GeomUtil.IsEqual(p, polygon[0])) continue;
                    vec2 = GeomUtil.SubXYZ(p, polygon[0]);
                }
                if (GeomUtil.IsOppositeDirection(vec1, vec2)) return PointComparePolygonResult.Boundary;
            }
            if (check1) return PointComparePolygonResult.Inside;
            return PointComparePolygonResult.Outside;
        }
        public static Line ConvertLine(Curve c)
        {
            return Line.CreateBound(c.GetEndPoint(0), c.GetEndPoint(1));
        }
        public static XYZ GetDirection(Curve c)
        {
            return GeomUtil.UnitVector(GeomUtil.SubXYZ(c.GetEndPoint(1), c.GetEndPoint(0)));
        }
        public static void CreateDetailLine(Curve c, Document doc, View v)
        {
            DetailLine dl = doc.Create.NewDetailCurve(v, c) as DetailLine;
        }
        public static void CreateDetailLinePolygon(Polygon pl, Document doc, View v)
        {
            foreach (Curve c in pl.ListCurve)
            {
                CreateDetailLine(c, doc, v);
            }
        }

        public static bool CreateListPolygon(List<Curve> listCurve, out List<Polygon> pls)
        {
            pls = new List<Polygon>();
            foreach (Curve c in listCurve)
            {
                List<Curve> cs = new List<Curve>();
                cs.Add(Line.CreateBound(c.GetEndPoint(0), c.GetEndPoint(1)));
                int i = 0; bool check = true;
                while (!GeomUtil.IsEqual(cs[0].GetEndPoint(0), cs[cs.Count - 1].GetEndPoint(1)))
                {
                    i++;
                    foreach (Curve c1 in listCurve)
                    {
                        XYZ pnt = cs[cs.Count - 1].GetEndPoint(1);
                        XYZ prePnt = cs[cs.Count - 1].GetEndPoint(0);
                        if (GeomUtil.IsEqual(pnt, c1.GetEndPoint(0)))
                        {
                            if (GeomUtil.IsEqual(prePnt, c1.GetEndPoint(1)))
                            {
                                continue;
                            }
                            cs.Add(Line.CreateBound(c1.GetEndPoint(0), c1.GetEndPoint(1)));
                            break;
                        }
                        else if (GeomUtil.IsEqual(pnt, c1.GetEndPoint(1)))
                        {
                            if (GeomUtil.IsEqual(prePnt, c1.GetEndPoint(0)))
                            {
                                continue;
                            }
                            cs.Add(Line.CreateBound(c1.GetEndPoint(1), c1.GetEndPoint(0)));
                            break;
                        }
                        else continue;
                    }
                    if (i == 200) { check = false; break; }
                }
                if (check)
                {
                    Polygon plgon = new Polygon(cs);

                    if (pls.Count == 0) pls.Add(plgon);
                    else
                    {
                        check = true;
                        foreach (Polygon pl in pls)
                        {
                            if (pl == plgon) { check = false; break; }
                        }
                        if (check) pls.Add(plgon);
                    }
                }
            }
            if (pls.Count == 0) return false;
            return true;
        }
        public static Polygon GetPolygonFromFaceFamilyInstance(FamilyInstance fi)
        {
            GeometryElement geoElem = fi.get_Geometry(new Options { ComputeReferences = true });
            List<Curve> cs = new List<Curve>();
            foreach (GeometryObject geoObj in geoElem)
            {
                GeometryInstance geoIns = geoObj as GeometryInstance;
                if (geoIns == null) continue;
                Transform tf = geoIns.Transform;
                foreach (GeometryObject geoSymObj in geoIns.GetSymbolGeometry())
                {
                    Curve c = geoSymObj as Line;
                    if (c != null)
                        cs.Add(GeomUtil.TransformCurve(c, tf));
                }
            }
            if (cs.Count < 3) throw new Exception("Incorrect input curve!");
            return new Polygon(cs);
        }
        public static XYZ ConvertStringToXYZ(string pointString)
        {
            List<double> nums = new List<double>();
            foreach (string s in pointString.Split('(', ',', ' ', ')'))
            {
                double x = 0;
                if (double.TryParse(s, out x)) { nums.Add(x); }
            }
            return new XYZ(nums[0], nums[1], nums[2]);
        }
        public static ViewSection CreateWallSection(Document linkedDoc, Document doc, ElementId id, string viewName, double offset)
        {
            Element e = linkedDoc.GetElement(id);
            if (!(e is Wall)) throw new Exception("Element is not a wall!");
            Wall wall = (Wall)e;
            Line line = (wall.Location as LocationCurve).Curve as Line;

            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);

            XYZ p1 = line.GetEndPoint(0), p2 = line.GetEndPoint(1);
            List<XYZ> ps = new List<XYZ> { p1, p2 }; ps.Sort(new ZYXComparer());
            p1 = ps[0]; p2 = ps[1];

            BoundingBoxXYZ bb = wall.get_BoundingBox(null);
            double minZ = bb.Min.Z, maxZ = bb.Max.Z;

            double l = GeomUtil.GetLength(GeomUtil.SubXYZ(p2, p1));
            double h = maxZ - minZ;
            double w = wall.WallType.Width;

            XYZ min = new XYZ(-l / 2 - offset, minZ - offset, -w - offset);
            XYZ max = new XYZ(l / 2 + offset, maxZ + offset, w + offset);

            Transform tf = Transform.Identity;
            tf.Origin = (p1 + p2) / 2;
            tf.BasisX = GeomUtil.UnitVector(p1 - p2);
            tf.BasisY = XYZ.BasisZ;
            tf.BasisZ = GeomUtil.CrossMatrix(tf.BasisX, tf.BasisY);

            BoundingBoxXYZ sectionBox = new BoundingBoxXYZ() { Transform = tf, Min = min, Max = max };
            ViewSection vs = ViewSection.CreateSection(doc, vft.Id, sectionBox);

            XYZ wallDir = GeomUtil.UnitVector(p2 - p1);
            XYZ upDir = XYZ.BasisZ;
            XYZ viewDir = GeomUtil.CrossMatrix(wallDir, upDir);

            min = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(p1, -wallDir, offset), -viewDir, offset);
            min = new XYZ(min.X, min.Y, minZ - offset);
            max = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(p2, wallDir, offset), viewDir, offset);
            max = new XYZ(max.X, max.Y, maxZ + offset);

            tf = vs.get_BoundingBox(null).Transform.Inverse;
            max = tf.OfPoint(max);
            min = tf.OfPoint(min);
            double maxx = 0, maxy = 0, maxz = 0, minx = 0, miny = 0, minz = 0;
            if (max.Z > min.Z)
            {
                maxz = max.Z;
                minz = min.Z;
            }
            else
            {
                maxz = min.Z;
                minz = max.Z;
            }


            if (Math.Round(max.X, 4) == Math.Round(min.X, 4))
            {
                maxx = max.X;
                minx = minz;
            }
            else if (max.X > min.X)
            {
                maxx = max.X;
                minx = min.X;
            }

            else
            {
                maxx = min.X;
                minx = max.X;
            }

            if (Math.Round(max.Y, 4) == Math.Round(min.Y, 4))
            {
                maxy = max.Y;
                miny = minz;
            }
            else if (max.Y > min.Y)
            {
                maxy = max.Y;
                miny = min.Y;
            }

            else
            {
                maxy = min.Y;
                miny = max.Y;
            }

            BoundingBoxXYZ sectionView = new BoundingBoxXYZ();
            sectionView.Max = new XYZ(maxx, maxy, maxz);
            sectionView.Min = new XYZ(minx, miny, minz);

            vs.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId);

            vs.get_Parameter(BuiltInParameter.VIEWER_BOUND_FAR_CLIPPING).Set(0);

            vs.CropBoxActive = true;
            vs.CropBoxVisible = true;

            doc.Regenerate();

            vs.CropBox = sectionView;
            vs.Name = viewName;
            return vs;
        }
        public static ViewSection CreateWallSection(Document linkedDoc, Document doc, Polygon directPolygon, ElementId id, string viewName, double offset)
        {
            Element e = linkedDoc.GetElement(id);
            if (!(e is Wall)) throw new Exception("Element is not a wall!");
            Wall wall = (Wall)e;
            Line line = (wall.Location as LocationCurve).Curve as Line;

            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);

            XYZ p1 = line.GetEndPoint(0), p2 = line.GetEndPoint(1);
            List<XYZ> ps = new List<XYZ> { p1, p2 }; ps.Sort(new ZYXComparer());
            p1 = ps[0]; p2 = ps[1];

            BoundingBoxXYZ bb = wall.get_BoundingBox(null);
            double minZ = bb.Min.Z, maxZ = bb.Max.Z;

            double l = GeomUtil.GetLength(GeomUtil.SubXYZ(p2, p1));
            double h = maxZ - minZ;
            double w = wall.WallType.Width;

            XYZ tfMin = new XYZ(-l / 2 - offset, minZ - offset, -w - offset);
            XYZ tfMax = new XYZ(l / 2 + offset, maxZ + offset, w + offset);

            XYZ wallDir = GeomUtil.UnitVector(p2 - p1);
            XYZ upDir = XYZ.BasisZ;
            XYZ viewDir = GeomUtil.CrossMatrix(wallDir, upDir);

            XYZ midPoint = (p1 + p2) / 2;
            XYZ pMidPoint = GetProjectPoint(directPolygon.Plane, midPoint);

            XYZ pPnt = GeomUtil.OffsetPoint(pMidPoint, viewDir, w * 10);
            if (GeomUtil.IsBigger(GeomUtil.GetLength(pMidPoint, directPolygon.CentralXYZPoint), GeomUtil.GetLength(pPnt, directPolygon.CentralXYZPoint)))
            {
                wallDir = -wallDir;
                upDir = XYZ.BasisZ;
                viewDir = GeomUtil.CrossMatrix(wallDir, upDir);
            }
            else
            {

            }

            pPnt = GeomUtil.OffsetPoint(p1, wallDir, offset);
            XYZ min = null, max = null;
            if (GeomUtil.IsBigger(GeomUtil.GetLength(GeomUtil.SubXYZ(pPnt, midPoint)), GeomUtil.GetLength(GeomUtil.SubXYZ(p1, midPoint))))
            {
                min = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(p1, wallDir, offset), -viewDir, offset);
                max = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(p2, -wallDir, offset), viewDir, offset);
            }
            else
            {
                min = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(p1, -wallDir, offset), -viewDir, offset);
                max = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(p2, wallDir, offset), viewDir, offset);
            }
            min = new XYZ(min.X, min.Y, minZ - offset);
            max = new XYZ(max.X, max.Y, maxZ + offset);

            Transform tf = Transform.Identity;
            tf.Origin = (p1 + p2) / 2;
            tf.BasisX = wallDir;
            tf.BasisY = XYZ.BasisZ;
            tf.BasisZ = GeomUtil.CrossMatrix(wallDir, upDir);

            BoundingBoxXYZ sectionBox = new BoundingBoxXYZ() { Transform = tf, Min = tfMin, Max = tfMax };
            ViewSection vs = ViewSection.CreateSection(doc, vft.Id, sectionBox);

            tf = vs.get_BoundingBox(null).Transform.Inverse;
            max = tf.OfPoint(max);
            min = tf.OfPoint(min);
            double maxx = 0, maxy = 0, maxz = 0, minx = 0, miny = 0, minz = 0;
            if (max.Z > min.Z)
            {
                maxz = max.Z;
                minz = min.Z;
            }
            else
            {
                maxz = min.Z;
                minz = max.Z;
            }


            if (Math.Round(max.X, 4) == Math.Round(min.X, 4))
            {
                maxx = max.X;
                minx = minz;
            }
            else if (max.X > min.X)
            {
                maxx = max.X;
                minx = min.X;
            }

            else
            {
                maxx = min.X;
                minx = max.X;
            }

            if (Math.Round(max.Y, 4) == Math.Round(min.Y, 4))
            {
                maxy = max.Y;
                miny = minz;
            }
            else if (max.Y > min.Y)
            {
                maxy = max.Y;
                miny = min.Y;
            }

            else
            {
                maxy = min.Y;
                miny = max.Y;
            }

            BoundingBoxXYZ sectionView = new BoundingBoxXYZ();
            sectionView.Max = new XYZ(maxx, maxy, maxz);
            sectionView.Min = new XYZ(minx, miny, minz);

            vs.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId);

            vs.get_Parameter(BuiltInParameter.VIEWER_BOUND_FAR_CLIPPING).Set(0);

            vs.CropBoxActive = true;
            vs.CropBoxVisible = true;

            doc.Regenerate();

            vs.CropBox = sectionView;
            vs.Name = viewName;
            return vs;
        }
        public static View CreateFloorCallout(Document doc, List<View> views, string level, BoundingBoxXYZ bb, string viewName, double offset)
        {
            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.FloorPlan == x.ViewFamily);
            XYZ max = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(bb.Max, XYZ.BasisX, offset), XYZ.BasisY, offset), XYZ.BasisZ, offset);
            XYZ min = GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(GeomUtil.OffsetPoint(bb.Min, -XYZ.BasisX, offset), -XYZ.BasisY, offset), -XYZ.BasisZ, offset);
            bb = new BoundingBoxXYZ { Max = max, Min = min };
            View pv = null;
            string s = string.Empty;
            bool check = false;
            foreach (View v in views)
            {
                try
                {
                    s = v.LookupParameter("Associated Level").AsString();
                    if (s == level) { pv = v; check = true; break; }
                }
                catch
                {
                    continue;
                }
            }
            if (!check) throw new Exception("Invalid level name!");
            View vs = ViewSection.CreateCallout(doc, pv.Id, vft.Id, min, max);
            vs.CropBox = bb;
            vs.Name = viewName;
            return vs;
        }
        public static string GetDirectoryPath(Document doc)
        {
            return Path.GetDirectoryName(doc.PathName);
        }
        public static string GetDirectoryPath(string documentName)
        {
            return Path.GetDirectoryName(documentName);
        }
        public static string CreateNameWithDocumentPathName(Document doc, string name, string exten)
        {
            string s = GetDirectoryPath(doc);
            string s1 = doc.PathName.Substring(s.Length + 1);
            return Path.Combine(s, s1.Substring(0, s1.Length - 4) + name + "." + exten);
        }
        public static string CreateNameWithDocumentPathName(string documentName, string name, string exten)
        {
            string s = GetDirectoryPath(documentName);
            string s1 = documentName.Substring(s.Length + 1);
            return Path.Combine(s, s1.Substring(0, s1.Length - 4) + name + "." + exten);
        }
        public static bool IsFileInUse(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("'path' cannot be null or empty.", "path");

            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read)) { }
            }
            catch (IOException)
            {
                return true;
            }

            return false;
        }
        public static string ConvertBoundingBoxToString(BoundingBoxXYZ bb)
        {
            return "{" + bb.Min.ToString() + ";" + bb.Max.ToString() + "}";
        }
        public static BoundingBoxXYZ ConvertStringToBoundingBox(string bbString)
        {
            BoundingBoxXYZ bb = new BoundingBoxXYZ();
            string[] ss = bbString.Split(';');
            ss[0] = ss[0].Substring(1, ss[0].Length - 1); ss[1] = ss[1].Substring(0, ss[1].Length - 1);
            bb.Min = ConvertStringToXYZ(ss[0]);
            bb.Max = ConvertStringToXYZ(ss[1]);
            return bb;
        }
        public static string ConvertPolygonToString(Polygon plgon)
        {
            string s = "{";
            for (int i = 0; i < plgon.ListXYZPoint.Count; i++)
            {
                if (i != plgon.ListXYZPoint.Count - 1)
                {
                    s += plgon.ListXYZPoint[i].ToString() + ";";
                }
                else
                {
                    s += plgon.ListXYZPoint[i].ToString() + "}";
                }
            }
            return s;
        }
        public static Polygon ConvertStringToPolygon(string bbString)
        {
            BoundingBoxXYZ bb = new BoundingBoxXYZ();
            string[] ss = bbString.Split(';');
            ss[0] = ss[0].Substring(1, ss[0].Length - 1);
            ss[ss.Length - 1] = ss[ss.Length - 1].Substring(0, ss[ss.Length - 1].Length - 1);
            List<XYZ> points = new List<XYZ>();
            foreach (string s in ss)
            {
                points.Add(ConvertStringToXYZ(s));
            }
            return new Polygon(points);
        }
        public static List<Curve> ConvertCurveLoopToCurveList(CurveLoop cl)
        {
            List<Curve> cs = new List<Curve>();
            foreach (Curve c in cl)
            {
                cs.Add(c);
            }
            return cs;
        }
        public static Polygon GetPolygonCut(Polygon mainPolygon, Polygon secPolygon)
        {
            PolygonComparePolygonResult res = new PolygonComparePolygonResult(mainPolygon, secPolygon);
            if (res.ListPolygon[0] != secPolygon) throw new Exception("Secondary Polygon must inside Main Polygon!");
            bool isInside = true;
            List<Curve> cs = new List<Curve>();
            foreach (Curve c in secPolygon.ListCurve)
            {
                LineComparePolygonResult lpRes = new LineComparePolygonResult(mainPolygon, CheckGeometry.ConvertLine(c));
                if (lpRes.Type == LineComparePolygonType.Inside)
                {
                    foreach (Curve c1 in mainPolygon.ListCurve)
                    {
                        LineCompareLineResult llres = new LineCompareLineResult(c, c1);
                        if (llres.Type == LineCompareLineType.SameDirectionLineOverlap)
                        {
                            goto Here;
                        }
                    }
                    cs.Add(c);
                }
                Here: continue;
            }
            isInside = false;
            foreach (Curve c in mainPolygon.ListCurve)
            {
                LineComparePolygonResult lpRes = new LineComparePolygonResult(secPolygon, CheckGeometry.ConvertLine(c));
                if (lpRes.Type == LineComparePolygonType.Outside)
                {
                    cs.Add(c);
                    continue;
                }
                foreach (Curve c1 in secPolygon.ListCurve)
                {
                    LineCompareLineResult llRes = new LineCompareLineResult(c, c1);
                    if (llRes.Type == LineCompareLineType.SameDirectionLineOverlap)
                    {
                        isInside = false;
                        if (llRes.ListOuterLine.Count == 0) break;
                        foreach (Line l in llRes.ListOuterLine)
                        {
                            LineComparePolygonResult lpRes1 = new LineComparePolygonResult(secPolygon, l);
                            if (lpRes1.Type != LineComparePolygonType.Inside)
                                cs.Add(l);
                        }
                        break;
                    }
                }
            }
            if (isInside) throw new Exception("Secondary Polygon must be tangential with Main Polygon!");
            return new Polygon(cs);
        }
        public static List<Curve> GetCurvesCut(Polygon mainPolygon, Polygon secPolygon)
        {
            PolygonComparePolygonResult res = new PolygonComparePolygonResult(mainPolygon, secPolygon);
            if (res.ListPolygon[0] != secPolygon) throw new Exception("Secondary Polygon must inside Main Polygon!");
            bool isInside = true;
            List<Curve> cs = new List<Curve>();
            foreach (Curve c in secPolygon.ListCurve)
            {
                LineComparePolygonResult lpRes = new LineComparePolygonResult(mainPolygon, CheckGeometry.ConvertLine(c));
                if (lpRes.Type == LineComparePolygonType.Inside)
                {
                    foreach (Curve c1 in mainPolygon.ListCurve)
                    {
                        LineCompareLineResult llres = new LineCompareLineResult(c, c1);
                        if (llres.Type == LineCompareLineType.SameDirectionLineOverlap)
                        {
                            goto Here;
                        }
                    }
                    cs.Add(c);
                }
                Here: continue;
            }
            isInside = false;
            foreach (Curve c in mainPolygon.ListCurve)
            {
                LineComparePolygonResult lpRes = new LineComparePolygonResult(secPolygon, CheckGeometry.ConvertLine(c));
                if (lpRes.Type == LineComparePolygonType.Outside)
                {
                    cs.Add(c);
                    continue;
                }
                foreach (Curve c1 in secPolygon.ListCurve)
                {
                    LineCompareLineResult llRes = new LineCompareLineResult(c, c1);
                    if (llRes.Type == LineCompareLineType.SameDirectionLineOverlap)
                    {
                        isInside = false;
                        if (llRes.ListOuterLine.Count == 0) break;
                        foreach (Line l in llRes.ListOuterLine)
                        {
                            LineComparePolygonResult lpRes1 = new LineComparePolygonResult(secPolygon, l);
                            if (lpRes1.Type != LineComparePolygonType.Inside)
                                cs.Add(l);
                        }
                        break;
                    }
                }
            }
            if (isInside) throw new Exception("Secondary Polygon must be tangential with Main Polygon!");
            return cs;
        }
        public static Transform GetTransform(Element e)
        {
            GeometryElement geoEle = e.get_Geometry(new Options() { ComputeReferences = true });
            foreach (GeometryObject geoObj in geoEle)
            {
                GeometryInstance geoIns = geoObj as GeometryInstance;
                if (geoIns != null)
                {
                    if (geoIns.Transform != null)
                    {
                        return geoIns.Transform;
                    }
                }
            }
            return null;
        }
        public static void InsertDetailItem(string familyName, XYZ location, Document doc, Transaction tx, View v, params string[] property_Values)
        {
            Family f = null;
            FilteredElementCollector col = new FilteredElementCollector(doc).OfClass(typeof(Family));
            string s = string.Empty;
            bool check = false;
            foreach (Element e in col)
            {
                if (e.Name == familyName)
                {
                    f = (Family)e;
                    check = true;
                    break;
                }
            }
            if (!check)
            {
                string filePath = new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath;
                string directoryPath = Path.GetDirectoryName(filePath);
                string fullFamilyName = Path.Combine(directoryPath, familyName + ".rfa");
                doc.LoadFamily(fullFamilyName, out f);
            }

            FamilySymbol symbol = null;
            foreach (ElementId symbolId in f.GetFamilySymbolIds())
            {
                symbol = doc.GetElement(symbolId) as FamilySymbol;
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    //tx.Commit();
                    //tx.Start();
                }
                break;
            }
            FamilyInstance fi = doc.Create.NewFamilyInstance(location, symbol, v);

            for (int i = 0; i < property_Values.Length; i += 2)
            {
                fi.LookupParameter(property_Values[i]).Set(property_Values[i + 1]);
            }
        }
        public static void ShowTaskDialog(params string[] contents)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < contents.Length; i += 2)
            {
                if ((i + 1) <= contents.Length - 1)
                {
                    sb.Append(contents[i] + ":" + " \t" + contents[i + 1] + "\n");
                }
                else
                {
                    sb.Append(contents[i]);
                }
            }
            TaskDialog.Show("Revit", sb.ToString());
        }
        public static void CreateModelLinePolygon(Document doc, Polygon pl)
        {
            SketchPlane sp = SketchPlane.Create(doc, pl.Plane);
            foreach (Curve c in pl.ListCurve)
            {
                CreateModelLine(doc, sp, c);
            }
        }
        public static void CreateModelLine(Document doc, SketchPlane sp, Curve c)
        {
            if (sp != null)
            {
                doc.Create.NewModelCurve(c, sp);
                return;
            }
            if (GeomUtil.IsSameOrOppositeDirection(GetDirection(c), XYZ.BasisZ))
            {
                sp = SketchPlane.Create(doc, Plane.CreateByOriginAndBasis(c.GetEndPoint(0), XYZ.BasisZ, XYZ.BasisX));
                CreateModelLine(doc, sp, c);
                return;
            }
            XYZ vecY = GetDirection(c).CrossProduct(XYZ.BasisZ);
            sp = SketchPlane.Create(doc, Plane.CreateByOriginAndBasis(c.GetEndPoint(0), GeomUtil.UnitVector(GetDirection(c)), GeomUtil.UnitVector(vecY)));
            CreateModelLine(doc, sp, c);
            return;
        }
        public static Polygon OffsetPolygon(Polygon polygon, double distance, bool isInside = true)
        {
            List<XYZ> points = new List<XYZ>();
            for (int i = 0; i < polygon.ListXYZPoint.Count; i++)
            {
                XYZ vec = null;
                if (i == 0)
                {
                    Curve c1 = Line.CreateBound(polygon.ListXYZPoint[polygon.ListXYZPoint.Count - 1], polygon.ListXYZPoint[i]);
                    Curve c2 = Line.CreateBound(polygon.ListXYZPoint[i], polygon.ListXYZPoint[i + 1]);

                    vec = polygon.Normal.CrossProduct(GetDirection(c1)) + polygon.Normal.CrossProduct(GetDirection(c2));
                    //vec = polygon.Normal.CrossProduct(GetDirection(c2));
                }
                else if (i == polygon.ListXYZPoint.Count - 1)
                {
                    Curve c1 = Line.CreateBound(polygon.ListXYZPoint[i - 1], polygon.ListXYZPoint[i]);
                    Curve c2 = Line.CreateBound(polygon.ListXYZPoint[i], polygon.ListXYZPoint[0]);
                    vec = polygon.Normal.CrossProduct(GetDirection(c1)) + polygon.Normal.CrossProduct(GetDirection(c2));
                    //vec = polygon.Normal.CrossProduct(GetDirection(c2));
                }
                else
                {
                    Curve c1 = Line.CreateBound(polygon.ListXYZPoint[i - 1], polygon.ListXYZPoint[i]);
                    Curve c2 = Line.CreateBound(polygon.ListXYZPoint[i], polygon.ListXYZPoint[i + 1]);
                    vec = polygon.Normal.CrossProduct(GetDirection(c1)) + polygon.Normal.CrossProduct(GetDirection(c2));
                    //vec = polygon.Normal.CrossProduct(GetDirection(c2));
                }
                XYZ temp = null;
                if (isInside)
                    temp = GeomUtil.OffsetPoint(polygon.ListXYZPoint[i], vec, distance * Math.Sqrt(2));
                else
                    temp = GeomUtil.OffsetPoint(polygon.ListXYZPoint[i], -vec, distance * Math.Sqrt(2));
                //temp = (polygon.CheckXYZPointPosition(temp) == PointComparePolygonResult.Inside) ? temp : GeomUtil.OffsetPoint(polygon.ListXYZPoint[i], -vec, distance);
                points.Add(temp);
            }
            return new Polygon(points);
        }
    }

}
