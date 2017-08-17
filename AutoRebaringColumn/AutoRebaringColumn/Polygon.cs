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

namespace AutoRebaringColumn
{
    public class Polygon
    {
        public List<Curve> ListCurve { get; set; }
        public List<XYZ> ListXYZPoint { get; private set; }
        public List<UV> ListUVPoint { get; private set; }
        public XYZ XVector { get; private set; }
        public XYZ YVector { get; private set; }
        public XYZ XVecManual { get; private set; }
        public XYZ YVecManual { get; private set; }
        public Plane PlaneManual { get; private set; }
        public XYZ Normal { get; private set; }
        public XYZ Origin { get; private set; }
        public PlanarFace Face { get; set; }
        public Plane Plane { get; private set; }
        public XYZ CentralXYZPoint { get; private set; }
        public UV CentralUVPoint { get; private set; }
        public List<XYZ> TwoXYZPointsBoundary { get; private set; }
        public List<XYZ> TwoXYZPointsLimit { get; private set; }
        public List<UV> TwoUVPointsBoundary { get; private set; }
        public List<UV> TwoUVPointsLimit { get; private set; }
        public double Height { get; private set; }
        public double Width { get; private set; }
        public double Perimeter { get; private set; }
        public double Area { get; private set; }
        public Polygon(PlanarFace f)
        {
            this.Face = f;
            this.ListCurve = CheckGeometry.GetCurves(f);
            this.Plane = Plane.CreateByOriginAndBasis(Face.Origin, Face.XVector, Face.YVector);

            GetParameters();
        }
        public Polygon(List<Curve> cs)
        {
            this.ListCurve = new List<Curve>();
            int i = 0;
            ListCurve.Add(Line.CreateBound(cs[0].GetEndPoint(0), cs[0].GetEndPoint(1)));
            while (!GeomUtil.IsEqual(ListCurve[ListCurve.Count - 1].GetEndPoint(1), ListCurve[0].GetEndPoint(0)))
            {
                i++;
                foreach (Curve c in cs)
                {
                    XYZ pnt = ListCurve[ListCurve.Count - 1].GetEndPoint(1);
                    XYZ prePnt = ListCurve[ListCurve.Count - 1].GetEndPoint(0);
                    if (GeomUtil.IsEqual(pnt, c.GetEndPoint(0)))
                    {
                        if (GeomUtil.IsEqual(prePnt, c.GetEndPoint(1)))
                        {
                            continue;
                        }
                        ListCurve.Add(Line.CreateBound(c.GetEndPoint(0), c.GetEndPoint(1)));
                        break;
                    }
                    else if (GeomUtil.IsEqual(pnt, c.GetEndPoint(1)))
                    {
                        if (GeomUtil.IsEqual(prePnt, c.GetEndPoint(0)))
                        {
                            continue;
                        }
                        ListCurve.Add(Line.CreateBound(c.GetEndPoint(1), c.GetEndPoint(0)));
                        break;
                    }
                    else continue;
                }
                if (i == 200) throw new Exception("Error when creating polygon");
            }
            XYZ origin = ListCurve[0].GetEndPoint(0);
            XYZ vecX = GeomUtil.UnitVector(CheckGeometry.GetDirection(cs[0]));
            XYZ vecT = GeomUtil.UnitVector(CheckGeometry.GetDirection(ListCurve[ListCurve.Count - 1]));
            XYZ normal = GeomUtil.UnitVector(GeomUtil.CrossMatrix(vecX, vecT));
            XYZ vecY = GeomUtil.UnitVector(GeomUtil.CrossMatrix(vecX, normal));
            this.Plane = Plane.CreateByOriginAndBasis(origin, vecX, vecY);

            GetParameters();
        }
        public Polygon(List<XYZ> points)
        {
            List<Curve> cs = new List<Curve>();
            for (int i = 0; i < points.Count; i++)
            {
                if (i < points.Count - 1)
                {
                    cs.Add(Line.CreateBound(points[i], points[i + 1]));
                }
                else
                {
                    cs.Add(Line.CreateBound(points[i], points[0]));
                }
            }

            this.ListCurve = cs;

            XYZ origin = cs[0].GetEndPoint(0);
            XYZ vecX = GeomUtil.UnitVector(CheckGeometry.GetDirection(cs[0]));
            XYZ vecT = GeomUtil.UnitVector(CheckGeometry.GetDirection(ListCurve[ListCurve.Count - 1]));
            XYZ normal = GeomUtil.UnitVector(GeomUtil.CrossMatrix(vecX, vecT));
            XYZ vecY = GeomUtil.UnitVector(GeomUtil.CrossMatrix(vecX, normal));
            this.Plane = Plane.CreateByOriginAndBasis(origin, vecX, vecY);

            GetParameters();
        }
        private void GetParameters()
        {
            List<XYZ> points = new List<XYZ>();
            foreach (Curve c in this.ListCurve)
            {
                points.Add(c.GetEndPoint(0));
            }
            this.ListXYZPoint = points;

            List<UV> uvpoints = new List<UV>();
            CentralXYZPoint = new XYZ(0, 0, 0);
            foreach (XYZ p in ListXYZPoint)
            {
                uvpoints.Add(CheckGeometry.Evaluate(this.Plane, p));
                CentralXYZPoint = GeomUtil.AddXYZ(CentralXYZPoint, p);
            }
            this.ListUVPoint = uvpoints;
            this.CentralXYZPoint = new XYZ(CentralXYZPoint.X / ListXYZPoint.Count, CentralXYZPoint.Y / ListXYZPoint.Count, CentralXYZPoint.Z / ListXYZPoint.Count);
            this.CentralUVPoint = CheckGeometry.Evaluate(this.Plane, CentralXYZPoint);
            this.XVector = this.Plane.XVec; this.YVector = this.Plane.YVec; this.Normal = this.Plane.Normal; this.Origin = this.Plane.Origin;

            GetPerimeter(); GetArea();
        }
        private void GetArea()
        {
            int j;
            double area = 0;

            for (int i = 0; i < ListUVPoint.Count; i++)
            {
                j = (i + 1) % ListUVPoint.Count;

                area += ListUVPoint[i].U * ListUVPoint[j].V;
                area -= ListUVPoint[i].V * ListUVPoint[j].U;
            }

            area /= 2;
            this.Area = (area < 0 ? -area : area);
        }
        private void GetPerimeter()
        {
            double len = 0;
            foreach (Curve c in ListCurve)
            {
                len += GeomUtil.GetLength(CheckGeometry.ConvertLine(c));
            }
            this.Perimeter = len;
        }
        public void SetManualDirection(XYZ vec, bool isXVector = true)
        {
            if (!GeomUtil.IsEqual(GeomUtil.DotMatrix(vec, Normal), 0)) throw new Exception("Input vector is not perpendicular with Normal!");
            XYZ xvec = null, yvec = null;
            if (isXVector)
            {
                xvec = GeomUtil.UnitVector(vec);
                yvec = GeomUtil.UnitVector(GeomUtil.CrossMatrix(xvec, this.Normal));
            }
            else
            {
                yvec = GeomUtil.UnitVector(vec);
                xvec = GeomUtil.UnitVector(GeomUtil.CrossMatrix(yvec, this.Normal));
            }
            this.XVecManual = GeomUtil.IsBigger(xvec, -xvec) ? xvec : -xvec;
            this.YVecManual = GeomUtil.IsBigger(yvec, -yvec) ? yvec : -yvec;
            this.PlaneManual = Plane.CreateByOriginAndBasis(CentralXYZPoint, XVecManual, YVecManual);
        }
        public void SetTwoPointsBoundary(XYZ vec, bool isXVector = true)
        {
            SetManualDirection(vec, isXVector);
            double maxU = 0, maxV = 0;
            foreach (XYZ xyzP in ListXYZPoint)
            {
                UV uvP = CheckGeometry.Evaluate(this.PlaneManual, xyzP);
                if (GeomUtil.IsBigger(Math.Abs(uvP.U), maxU)) maxU = Math.Abs(uvP.U);
                if (GeomUtil.IsBigger(Math.Abs(uvP.V), maxV)) maxV = Math.Abs(uvP.V);
            }
            UV uvboundP = new UV(-maxU, -maxV);
            XYZ p1 = CheckGeometry.Evaluate(this.PlaneManual, uvboundP), p2 = CheckGeometry.Evaluate(this.PlaneManual, -uvboundP);
            TwoXYZPointsBoundary = new List<XYZ> { p1, p2 };
            TwoUVPointsBoundary = new List<UV> { CheckGeometry.Evaluate(this.Plane, p1), CheckGeometry.Evaluate(this.Plane, p2) };
        }
        public void SetTwoPointsLimit(XYZ vec, bool isXVector = true)
        {
            SetManualDirection(vec, isXVector);
            double maxU = 0, maxV = 0, minU = 0, minV = 0;
            foreach (XYZ xyzP in ListXYZPoint)
            {
                UV uvP = CheckGeometry.Evaluate(this.PlaneManual, xyzP);
                if (GeomUtil.IsBigger(uvP.U, maxU)) maxU = uvP.U;
                if (GeomUtil.IsBigger(uvP.V, maxV)) maxV = uvP.V;
                if (GeomUtil.IsSmaller(uvP.U, minU)) minU = uvP.U;
                if (GeomUtil.IsSmaller(uvP.V, minV)) minV = uvP.V;
            }
            UV min = new UV(minU, minV), max = new UV(maxU, maxV);
            TwoUVPointsLimit = new List<UV> { min, max };
            XYZ p1 = CheckGeometry.Evaluate(this.PlaneManual, min), p2 = CheckGeometry.Evaluate(this.PlaneManual, max);
            TwoXYZPointsLimit = new List<XYZ> { p1, p2 };
        }
        public void SetTwoDimension(XYZ vec, bool isXVector = true)
        {
            SetManualDirection(vec, isXVector);
            double maxU = 0, maxV = 0;
            List<UV> uvPs = new List<UV>();
            foreach (XYZ xyzP in ListXYZPoint)
            {
                uvPs.Add(CheckGeometry.Evaluate(this.PlaneManual, xyzP));
            }
            for (int i = 0; i < uvPs.Count; i++)
            {
                for (int j = i + 1; j < uvPs.Count; j++)
                {
                    if (GeomUtil.IsBigger(Math.Abs(uvPs[i].U - uvPs[j].U), maxU)) maxU = Math.Abs(uvPs[i].U - uvPs[j].U);
                    if (GeomUtil.IsBigger(Math.Abs(uvPs[i].V - uvPs[j].V), maxV)) maxV = Math.Abs(uvPs[i].V - uvPs[j].V);
                }
            }
            this.Width = maxU;
            this.Height = maxV;
        }
        public bool IsPointInPolygon(UV p)
        {
            List<UV> polygon = this.ListUVPoint;
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
        public bool IsPointInPolygonNewCheck(UV p)
        {
            List<UV> polygon = this.ListUVPoint;
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

            if (!GeomUtil.IsBigger(p.U, minX) || !GeomUtil.IsSmaller(p.U, maxX) || !GeomUtil.IsBigger(p.V, minY) || !GeomUtil.IsSmaller(p.V, maxY))
            {
                return false;
            }
            bool inside = false;
            for (int i = 0, j = polygon.Count - 1; i < polygon.Count; j = i++)
            {
                if ((!GeomUtil.IsSmaller(polygon[i].V, p.V) != (!GeomUtil.IsSmaller(polygon[j].V, p.V)) &&
                     !GeomUtil.IsBigger(p.U, (polygon[j].U - polygon[i].U) * (p.V - polygon[i].V) / (polygon[j].V - polygon[i].V) + polygon[i].U)))
                {
                    inside = !inside;
                }
            }
            return inside;
        }
        public bool IsLineInPolygon(Line l)
        {
            XYZ p1 = l.GetEndPoint(0);
            if (CheckXYZPointPosition(l.GetEndPoint(0)) == PointComparePolygonResult.NonPlanar || CheckXYZPointPosition(l.GetEndPoint(0)) == PointComparePolygonResult.NonPlanar) return false;
            double len = GeomUtil.GetLength(l);
            XYZ dir = GeomUtil.UnitVector(l.Direction);
            if (GeomUtil.IsEqual(l.GetEndPoint(1), GeomUtil.OffsetPoint(p1, dir, len)))
            {
            }
            else if (GeomUtil.IsEqual(l.GetEndPoint(1), GeomUtil.OffsetPoint(p1, -dir, len)))
            {
                dir = -dir;
            }
            else throw new Exception("Error when retrieve result!");

            for (int i = 0; i <= 100; i++)
            {
                XYZ p = GeomUtil.OffsetPoint(p1, dir, len / 100 * i);
                if (CheckXYZPointPosition(p) == PointComparePolygonResult.Outside) return false;
            }
            return true;
        }
        public bool IsLineOutPolygon(Line l)
        {
            XYZ p1 = l.GetEndPoint(0);
            if (CheckXYZPointPosition(l.GetEndPoint(0)) == PointComparePolygonResult.NonPlanar || CheckXYZPointPosition(l.GetEndPoint(0)) == PointComparePolygonResult.NonPlanar) return false;
            double len = GeomUtil.GetLength(l);
            XYZ vec = GeomUtil.IsEqual(GeomUtil.AddXYZ(p1, GeomUtil.MultiplyVector(l.Direction, len)), l.GetEndPoint(1)) ? l.Direction : -l.Direction;
            int count = 0;
            for (int i = 0; i <= 100; i++)
            {
                XYZ p = GeomUtil.OffsetPoint(p1, vec, len / 100 * i);
                if (CheckXYZPointPosition(p) != PointComparePolygonResult.Outside)
                {
                    if (CheckXYZPointPosition(p) == PointComparePolygonResult.Inside) return false;
                    count++;
                }
            }
            if (count > 2) return false;
            return true;
        }
        public PointComparePolygonResult CheckUVPointPosition(UV p)
        {
            List<UV> polygon = this.ListUVPoint;
            bool check1 = IsPointInPolygon(p);
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
        public PointComparePolygonResult CheckXYZPointPosition(XYZ p)
        {
            if (!GeomUtil.IsEqual(CheckGeometry.GetSignedDistance(this.Plane, p), 0)) return PointComparePolygonResult.NonPlanar;
            UV uvP = Evaluate(p);
            return CheckUVPointPosition(uvP);
        }
        public XYZ Evaluate(UV p) { return CheckGeometry.Evaluate(this.Plane, p); }
        public UV Evaluate(XYZ p) { return CheckGeometry.Evaluate(this.Plane, p); }
        public XYZ GetTopDirectionFromCurve()
        {
            List<XYZ> vecs = new List<XYZ>();
            foreach (Curve c in this.ListCurve)
            {
                XYZ vec = GeomUtil.UnitVector(CheckGeometry.GetDirection(c));
                vec = GeomUtil.IsBigger(vec, -vec) ? vec : -vec;
                vecs.Add(vec);
            }
            vecs.Sort(new ZYXComparer());
            return vecs[vecs.Count - 1];
        }
        public void OffsetPolygon(XYZ direction, double distance)
        {
            for (int i = 0; i < ListCurve.Count; i++)
            {
                ListCurve[i] = GeomUtil.OffsetCurve(ListCurve[i], direction, distance);
            }
        }
        public static bool operator ==(Polygon pl1, Polygon pl2)
        {
            try
            {
                List<XYZ> points = pl1.ListXYZPoint;
            }
            catch
            {
                try
                {
                    List<XYZ> points = pl2.ListXYZPoint;
                    return false;
                }
                catch
                {
                    return true;
                }
            }
            try
            {
                List<XYZ> points = pl2.ListXYZPoint;
            }
            catch
            {
                return false;
            }
            List<XYZ> pnts1 = pl1.ListXYZPoint, pnts2 = pl2.ListXYZPoint;
            pnts1.Sort(new ZYXComparer()); pnts2.Sort(new ZYXComparer());
            for (int i = 0; i < pnts1.Count; i++)
            {
                if (!GeomUtil.IsEqual(pnts1[i], pnts2[i])) return false;
            }
            return true;
        }
        public static bool operator !=(Polygon pl1, Polygon pl2)
        {
            return !(pl1 == pl2);
        }
        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
    public enum PointComparePolygonResult
    {
        Inside, Outside, Boundary, Node, NonPlanar
    }
    public class LineCompareLineResult
    {
        public LineCompareLineType Type { get; private set; }
        public Line Line { get; private set; }
        public List<Line> ListOuterLine { get; private set; }
        public Line MergeLine { get; private set; }
        public XYZ Point { get; private set; }
        private Line line1;
        private Line line2;
        public LineCompareLineResult(Line l1, Line l2)
        {
            this.line1 = l1; this.line2 = l2; GetParameter();
        }
        public LineCompareLineResult(Curve l1, Curve l2)
        {
            this.line1 = Line.CreateBound(l1.GetEndPoint(0), l1.GetEndPoint(1));
            this.line2 = Line.CreateBound(l2.GetEndPoint(0), l2.GetEndPoint(1));
            GetParameter();
        }
        private void GetParameter()
        {
            XYZ vec1 = line1.Direction, vec2 = line2.Direction;
            if (GeomUtil.IsSameOrOppositeDirection(vec1, vec2))
            {
                #region SameDirection
                if (GeomUtil.IsEqual(CheckGeometry.GetSignedDistance(line1, line2.GetEndPoint(0)), 0))
                {
                    if (GeomUtil.IsEqual(line1.GetEndPoint(0), line2.GetEndPoint(0)))
                    {
                        if (GeomUtil.IsOppositeDirection(GeomUtil.SubXYZ(line1.GetEndPoint(1), line1.GetEndPoint(0)), GeomUtil.SubXYZ(line2.GetEndPoint(1), line2.GetEndPoint(0))))
                        {
                            this.Point = line1.GetEndPoint(0);
                            this.MergeLine = Line.CreateBound(line1.GetEndPoint(1), line2.GetEndPoint(1));
                            this.Type = LineCompareLineType.SameDirectionPointOverlap; return;
                        }
                        if (CheckGeometry.IsPointInLine(line2, line1.GetEndPoint(1)))
                        {
                            this.Line = line1;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        }
                        else
                        {
                            this.Line = line2;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        }
                        ListOuterLine = new List<Line>();
                        try
                        {
                            Line l = Line.CreateBound(line1.GetEndPoint(1), line2.GetEndPoint(1));
                            ListOuterLine = new List<Line>() { l };
                        }
                        catch
                        { }
                        return;
                    }
                    else if (GeomUtil.IsEqual(line1.GetEndPoint(1), line2.GetEndPoint(0)))
                    {
                        if (GeomUtil.IsOppositeDirection(GeomUtil.SubXYZ(line1.GetEndPoint(0), line1.GetEndPoint(1)), GeomUtil.SubXYZ(line2.GetEndPoint(1), line2.GetEndPoint(0))))
                        {
                            this.Point = line1.GetEndPoint(1);
                            this.MergeLine = Line.CreateBound(line1.GetEndPoint(0), line2.GetEndPoint(1));
                            this.Type = LineCompareLineType.SameDirectionPointOverlap; return;
                        }
                        if (CheckGeometry.IsPointInLine(line2, line1.GetEndPoint(0)))
                        {
                            this.Line = line1;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        }
                        else
                        {
                            this.Line = line2;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        }
                        ListOuterLine = new List<Line>();
                        try
                        {
                            Line l = Line.CreateBound(line1.GetEndPoint(0), line2.GetEndPoint(1));
                            ListOuterLine = new List<Line>() { l };
                        }
                        catch
                        { }
                        return;
                    }
                    else if (GeomUtil.IsEqual(line1.GetEndPoint(1), line2.GetEndPoint(1)))
                    {
                        if (GeomUtil.IsOppositeDirection(GeomUtil.SubXYZ(line1.GetEndPoint(1), line1.GetEndPoint(0)), GeomUtil.SubXYZ(line2.GetEndPoint(1), line2.GetEndPoint(0))))
                        {
                            this.Point = line1.GetEndPoint(1);
                            this.MergeLine = Line.CreateBound(line1.GetEndPoint(0), line2.GetEndPoint(0));
                            this.Type = LineCompareLineType.SameDirectionPointOverlap; return;
                        }
                        if (CheckGeometry.IsPointInLine(line2, line1.GetEndPoint(0)))
                        {
                            this.Line = line1;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        }
                        else
                        {
                            this.Line = line2;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        }
                        ListOuterLine = new List<Line>();
                        try
                        {
                            Line l = Line.CreateBound(line1.GetEndPoint(0), line2.GetEndPoint(0));
                            ListOuterLine = new List<Line>() { l };
                        }
                        catch
                        { }
                        return;
                    }
                    else if (GeomUtil.IsEqual(line1.GetEndPoint(0), line2.GetEndPoint(1)))
                    {
                        if (GeomUtil.IsOppositeDirection(GeomUtil.SubXYZ(line1.GetEndPoint(0), line1.GetEndPoint(1)), GeomUtil.SubXYZ(line2.GetEndPoint(1), line2.GetEndPoint(0))))
                        {
                            this.Point = line1.GetEndPoint(0);
                            this.MergeLine = Line.CreateBound(line1.GetEndPoint(1), line2.GetEndPoint(0));
                            this.Type = LineCompareLineType.SameDirectionPointOverlap; return;
                        }
                        if (CheckGeometry.IsPointInLine(line2, line1.GetEndPoint(1)))
                        {
                            this.Line = line1;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        }
                        else
                        {
                            this.Line = line2;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        }
                        ListOuterLine = new List<Line>();
                        try
                        {
                            Line l = Line.CreateBound(line1.GetEndPoint(1), line2.GetEndPoint(0));
                            ListOuterLine = new List<Line>() { l };
                        }
                        catch
                        { }
                        return;
                    }
                    if (CheckGeometry.IsPointInLine(line2, line1.GetEndPoint(0)))
                    {
                        if (CheckGeometry.IsPointInLine(line2, line1.GetEndPoint(1)))
                        {
                            this.Line = line1;
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                            if (CheckGeometry.IsPointInLine(Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(1)), line1.GetEndPoint(0)))
                            {
                                Line l1 = Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(0));
                                Line l2 = Line.CreateBound(line2.GetEndPoint(1), line1.GetEndPoint(1));
                                ListOuterLine = new List<Line> { l1, l2 };
                            }
                            else
                            {
                                Line l1 = Line.CreateBound(line2.GetEndPoint(1), line1.GetEndPoint(0));
                                Line l2 = Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(1));
                                ListOuterLine = new List<Line> { l1, l2 };

                            }
                            return;
                        }
                        if (CheckGeometry.IsPointInLine(line1, line2.GetEndPoint(0)))
                        {
                            this.Line = Line.CreateBound(line1.GetEndPoint(0), line2.GetEndPoint(0));
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                            Line l1 = Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(1));
                            Line l2 = Line.CreateBound(line2.GetEndPoint(1), line1.GetEndPoint(0));
                            ListOuterLine = new List<Line> { l1, l2 };
                        }
                        else
                        {
                            this.Line = Line.CreateBound(line1.GetEndPoint(0), line2.GetEndPoint(1));
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                            Line l1 = Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(0));
                            Line l2 = Line.CreateBound(line2.GetEndPoint(1), line1.GetEndPoint(1));
                            ListOuterLine = new List<Line> { l1, l2 };
                        }
                        return;
                    }
                    if (CheckGeometry.IsPointInLine(line2, line1.GetEndPoint(1)))
                    {
                        if (CheckGeometry.IsPointInLine(line1, line2.GetEndPoint(0)))
                        {
                            this.Line = Line.CreateBound(line1.GetEndPoint(1), line2.GetEndPoint(0));
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                            Line l1 = Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(0));
                            Line l2 = Line.CreateBound(line2.GetEndPoint(1), line1.GetEndPoint(1));
                            ListOuterLine = new List<Line> { l1, l2 };
                        }
                        else
                        {
                            this.Line = Line.CreateBound(line1.GetEndPoint(1), line2.GetEndPoint(1));
                            this.Type = LineCompareLineType.SameDirectionLineOverlap;
                            Line l1 = Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(1));
                            Line l2 = Line.CreateBound(line2.GetEndPoint(1), line1.GetEndPoint(0));
                            ListOuterLine = new List<Line> { l1, l2 };
                        }
                        return;
                    }
                    if (CheckGeometry.IsPointInLine(line1, line2.GetEndPoint(0)))
                    {
                        this.Line = line2;
                        this.Type = LineCompareLineType.SameDirectionLineOverlap;
                        if (CheckGeometry.IsPointInLine(Line.CreateBound(line1.GetEndPoint(0), line2.GetEndPoint(1)), line2.GetEndPoint(0)))
                        {
                            Line l1 = Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(0));
                            Line l2 = Line.CreateBound(line2.GetEndPoint(1), line1.GetEndPoint(1));
                            ListOuterLine = new List<Line> { l1, l2 };
                        }
                        else
                        {
                            Line l1 = Line.CreateBound(line2.GetEndPoint(1), line1.GetEndPoint(0));
                            Line l2 = Line.CreateBound(line2.GetEndPoint(0), line1.GetEndPoint(1));
                            ListOuterLine = new List<Line> { l1, l2 };

                        }
                        return;
                    }
                    this.Type = LineCompareLineType.SameDirectionNonOverlap; return;
                }
                else
                { this.Type = LineCompareLineType.Parallel; return; }
                #endregion
            }
            XYZ p1 = line1.GetEndPoint(0), p2 = line1.GetEndPoint(1), p3 = line2.GetEndPoint(0), p4 = line2.GetEndPoint(1);
            if (CheckGeometry.IsPointInLineOrExtend(line2, p1))
            {
                this.Point = p1;
                if (CheckGeometry.IsPointInLine(line2, p1)) { this.Type = LineCompareLineType.Intersect; return; }
                this.Type = LineCompareLineType.NonIntersectPlanar; return;
            }
            if (CheckGeometry.IsPointInLineOrExtend(line2, p2))
            {
                this.Point = p2;
                if (CheckGeometry.IsPointInLine(line2, p2)) { this.Type = LineCompareLineType.Intersect; return; }
                this.Type = LineCompareLineType.NonIntersectPlanar; return;
            }
            if (CheckGeometry.IsPointInLineOrExtend(line1, p3))
            {
                this.Point = p3;
                if (CheckGeometry.IsPointInLine(line1, p3)) { this.Type = LineCompareLineType.Intersect; return; }
                this.Type = LineCompareLineType.NonIntersectPlanar; return;
            }
            if (CheckGeometry.IsPointInLineOrExtend(line1, p4))
            {
                this.Point = p4;
                if (CheckGeometry.IsPointInLine(line1, p4)) { this.Type = LineCompareLineType.Intersect; return; }
                this.Type = LineCompareLineType.NonIntersectPlanar; return;
            }
            if (GeomUtil.IsEqual(GeomUtil.DotMatrix(GeomUtil.SubXYZ(p1, p3), GeomUtil.CrossMatrix(vec1, vec2)), 0))
            {
                double h1 = CheckGeometry.GetSignedDistance(line2, p1), h2 = CheckGeometry.GetSignedDistance(line2, p2);
                double deltaH = 0, L1 = 0;
                double L = GeomUtil.GetLength(p1, p2);
                XYZ pP1 = CheckGeometry.GetProjectPoint(line2, p1), pP2 = CheckGeometry.GetProjectPoint(line2, p2);
                if (GeomUtil.IsEqual(pP1, p1))
                {
                    this.Point = p1; this.Type = LineCompareLineType.Intersect; return;
                }
                if (GeomUtil.IsEqual(pP2, p2))
                {
                    this.Point = p2; this.Type = LineCompareLineType.Intersect; return;
                }
                XYZ tP1 = null, tP2 = null;
                if (GeomUtil.IsSameDirection(GeomUtil.SubXYZ(pP1, p1), GeomUtil.SubXYZ(pP2, p2)))
                {
                    deltaH = Math.Abs(h1 - h2);
                    L1 = L * h1 / deltaH;
                    tP1 = GeomUtil.OffsetPoint(p1, line1.Direction, L1); tP2 = GeomUtil.OffsetPoint(p1, line1.Direction, -L1);
                    if (CheckGeometry.IsPointInLineOrExtend(line2, tP1)) { this.Point = tP1; }
                    else if (CheckGeometry.IsPointInLineOrExtend(line2, tP2)) { this.Point = tP2; }
                    else
                    {
                        throw new Exception("Two points is not in line extend!");
                    }
                    this.Type = LineCompareLineType.NonIntersectPlanar; return;
                }

                deltaH = h1 + h2;
                L1 = L * h1 / deltaH;
                tP1 = GeomUtil.OffsetPoint(p1, line1.Direction, L1); tP2 = GeomUtil.OffsetPoint(p1, line1.Direction, -L1);
                if (CheckGeometry.IsPointInLineOrExtend(line2, tP1)) { this.Point = tP1; }
                else if (CheckGeometry.IsPointInLineOrExtend(line2, tP2)) { this.Point = tP2; }
                else { throw new Exception("Two points is not in line extend!"); }
                if (CheckGeometry.IsPointInLine(line2, this.Point) && CheckGeometry.IsPointInLine(line1, this.Point))
                {

                    this.Type = LineCompareLineType.Intersect; return;
                }
                this.Type = LineCompareLineType.NonIntersectPlanar; return;
            }
            this.Type = LineCompareLineType.NonIntersectNonPlanar; return;
        }
    }
    public enum LineCompareLineType
    {
        SameDirectionPointOverlap, SameDirectionNonOverlap, SameDirectionLineOverlap, Parallel, Intersect, NonIntersectPlanar, NonIntersectNonPlanar
    }
    public class LineComparePolygonResult
    {
        public LineComparePolygonType Type { get; private set; }
        public List<Line> ListLine { get; private set; }
        public Line ProjectLine { get; private set; }
        public List<XYZ> ListPoint { get; private set; }
        private Line line;
        private Polygon polygon;
        public LineComparePolygonResult(Polygon plgon, Line l)
        {
            this.line = l; this.polygon = plgon; GetParameter();
        }
        private void GetParameter()
        {
            #region Planar
            if (GeomUtil.IsEqual(GeomUtil.DotMatrix(line.Direction, polygon.Normal), 0))
            {
                if (polygon.CheckXYZPointPosition(line.GetEndPoint(0)) == PointComparePolygonResult.NonPlanar)
                {
                    XYZ p11 = line.GetEndPoint(0), p21 = line.GetEndPoint(1);
                    XYZ pP11 = CheckGeometry.GetProjectPoint(polygon.Plane, p11);
                    XYZ pP21 = CheckGeometry.GetProjectPoint(polygon.Plane, p21);
                    Line l11 = Line.CreateBound(pP11, pP21);
                    this.ProjectLine = l11;
                    this.Type = LineComparePolygonType.NonPlanarParallel; return;
                }
                if (polygon.IsLineOutPolygon(line))
                {
                    this.Type = LineComparePolygonType.Outside; return;
                }
                this.ListLine = new List<Line>();
                this.ListPoint = new List<XYZ>();
                if (polygon.IsLineInPolygon(line))
                {
                    this.ListLine.Add(line); this.Type = LineComparePolygonType.Inside; return;
                }
                foreach (Curve c in polygon.ListCurve)
                {
                    LineCompareLineResult res2 = new LineCompareLineResult(c, line);
                    if (res2.Type == LineCompareLineType.SameDirectionLineOverlap)
                    {
                        this.ListLine.Add(res2.Line);
                    }
                    if (res2.Type == LineCompareLineType.Intersect)
                    {

                        this.ListPoint.Add(res2.Point);
                    }
                }
                if (ListPoint.Count != 0)
                {
                    ListPoint.Sort(new ZYXComparer());
                    List<XYZ> points = new List<XYZ>();
                    for (int i = 0; i < ListPoint.Count; i++)
                    {
                        bool check = true;
                        for (int j = i + 1; j < ListPoint.Count; j++)
                        {
                            if (GeomUtil.IsEqual(ListPoint[i], ListPoint[j]))
                            {
                                check = false; break;
                            }
                        }
                        if (check) points.Add(ListPoint[i]);
                    }
                    ListPoint = points;
                    if (ListPoint.Count == 1)
                    {
                        ListPoint.Insert(0, line.GetEndPoint(0)); ListPoint.Add(line.GetEndPoint(1));
                    }
                    else
                    {
                        if (GeomUtil.IsEqual(ListPoint[0], line.GetEndPoint(0)))
                        {
                            if (GeomUtil.IsEqual(ListPoint[ListPoint.Count - 1], line.GetEndPoint(1)))
                            { }
                            else ListPoint.Add(line.GetEndPoint(1));
                        }
                        else if (GeomUtil.IsEqual(ListPoint[0], line.GetEndPoint(1)))
                        {
                            if (GeomUtil.IsEqual(ListPoint[ListPoint.Count - 1], line.GetEndPoint(0)))
                            { }
                            else ListPoint.Add(line.GetEndPoint(0));
                        }
                        else if (GeomUtil.IsEqual(ListPoint[ListPoint.Count - 1], line.GetEndPoint(0)))
                        {
                            ListPoint.Insert(0, line.GetEndPoint(1));
                        }
                        else if (GeomUtil.IsEqual(ListPoint[ListPoint.Count - 1], line.GetEndPoint(1)))
                        {
                            ListPoint.Insert(0, line.GetEndPoint(0));
                        }
                        else if (GeomUtil.IsSameDirection(GeomUtil.SubXYZ(ListPoint[ListPoint.Count - 1], ListPoint[0]), GeomUtil.SubXYZ(line.GetEndPoint(1), line.GetEndPoint(0))))
                        {
                            ListPoint.Insert(0, line.GetEndPoint(0)); ListPoint.Add(line.GetEndPoint(1));
                        }
                        else
                        {
                            ListPoint.Insert(0, line.GetEndPoint(1)); ListPoint.Add(line.GetEndPoint(0));
                        }
                    }
                    for (int i = 0; i < this.ListPoint.Count - 1; i++)
                    {
                        if (GeomUtil.IsEqual(ListPoint[i], ListPoint[i + 1])) continue;
                        Line l = null;
                        try
                        {
                            l = Line.CreateBound(ListPoint[i], ListPoint[i + 1]);
                        }
                        catch
                        {
                            continue;
                        }
                        bool check = true;
                        if (polygon.IsLineInPolygon(l))
                        {
                            bool check2 = false;
                            for (int j = 0; j < ListLine.Count; j++)
                            {
                                LineCompareLineResult res1 = new LineCompareLineResult(ListLine[j], l);
                                if (res1.Type == LineCompareLineType.SameDirectionLineOverlap) check = false;
                                if (res1.Type == LineCompareLineType.SameDirectionPointOverlap)
                                {
                                    ListLine[j] = res1.MergeLine;
                                    check2 = true; break;
                                }
                            }
                            if (check2) continue;
                            if (check)
                            {
                                ListLine.Add(l);
                            }
                        }
                    }
                    this.Type = LineComparePolygonType.OverlapOrIntersect; return;
                }
            }
            #endregion
            XYZ p1 = line.GetEndPoint(0), p2 = line.GetEndPoint(1);
            XYZ pP1 = CheckGeometry.GetProjectPoint(polygon.Plane, p1);
            XYZ pP2 = CheckGeometry.GetProjectPoint(polygon.Plane, p2);
            ListPoint = new List<XYZ>();
            if (GeomUtil.IsEqual(pP1, pP2))
            {
                this.ListPoint.Add(pP1);
                if (CheckGeometry.IsPointInLine(line, pP1))
                {
                    if (polygon.CheckXYZPointPosition(pP1) != PointComparePolygonResult.Outside) { this.Type = LineComparePolygonType.PerpendicularIntersectFace; return; }
                    this.Type = LineComparePolygonType.PerpendicularIntersectPlane; return;
                }
                this.Type = LineComparePolygonType.PerpendicularNonIntersect; return;
            }
            Line l1 = Line.CreateBound(pP1, pP2);
            ProjectLine = l1;
            LineCompareLineResult res = new LineCompareLineResult(line, l1);
            if (res.Type == LineCompareLineType.Intersect)
            {
                PointComparePolygonResult resP = polygon.CheckXYZPointPosition(res.Point);
                ListPoint.Add(res.Point);
                if (resP == PointComparePolygonResult.Outside) { this.Type = LineComparePolygonType.NonPlanarIntersectPlane; return; }
                this.Type = LineComparePolygonType.NonPlanarIntersectFace; return;
            }
            ListPoint.Add(res.Point);
            this.Type = LineComparePolygonType.NonPlanarNonIntersect; return;
        }
    }
    public enum LineComparePolygonType
    {
        NonPlanarIntersectPlane, NonPlanarIntersectFace, NonPlanarNonIntersect, NonPlanarParallel, Outside, Inside, OverlapOrIntersect,
        PerpendicularNonIntersect, PerpendicularIntersectPlane, PerpendicularIntersectFace
    }
    public class PolygonComparePolygonResult
    {
        public PolygonComparePolygonPositionType PositionType { get; private set; }
        public PolygonComparePolygonIntersectType IntersectType { get; private set; }
        public List<Line> ListLine { get; private set; }
        public List<XYZ> ListPoint { get; private set; }
        public MultiPolygon OuterMultiPolygon { get; private set; }
        public List<Polygon> ListPolygon { get; private set; }
        private Polygon polygon1;
        private Polygon polygon2;
        public PolygonComparePolygonResult(Polygon pl1, Polygon pl2)
        {
            this.polygon1 = pl1; this.polygon2 = pl2; GetPositionType(); GetIntersectTypeAndOtherParameter();
        }
        private void GetPositionType()
        {
            if (GeomUtil.IsSameOrOppositeDirection(polygon1.Normal, polygon2.Normal))
            {
                if (polygon1.CheckXYZPointPosition(polygon2.ListXYZPoint[0]) != PointComparePolygonResult.NonPlanar) { this.PositionType = PolygonComparePolygonPositionType.Planar; return; }
                this.PositionType = PolygonComparePolygonPositionType.Parallel; return;
            }
            this.PositionType = PolygonComparePolygonPositionType.NonPlanar; return;
        }
        private void GetIntersectTypeAndOtherParameter()
        {
            switch (PositionType)
            {
                case PolygonComparePolygonPositionType.Parallel: this.IntersectType = PolygonComparePolygonIntersectType.NonIntersect; return;
                #region NonPlanar
                case PolygonComparePolygonPositionType.NonPlanar:
                    bool check = false, check2 = false;
                    List<XYZ> points = new List<XYZ>();
                    List<Line> lines = new List<Line>();
                    foreach (Curve c in polygon2.ListCurve)
                    {
                        LineComparePolygonResult res = new LineComparePolygonResult(polygon1, CheckGeometry.ConvertLine(c));
                        if (res.Type == LineComparePolygonType.NonPlanarIntersectFace || res.Type == LineComparePolygonType.PerpendicularIntersectFace)
                        {
                            check = true;
                            bool checkP = true;
                            foreach (XYZ point in points)
                            {
                                if (GeomUtil.IsEqual(point, res.ListPoint[0])) { checkP = false; break; }
                            }
                            if (checkP)
                                points.Add(res.ListPoint[0]);
                        }
                        if (res.Type == LineComparePolygonType.OverlapOrIntersect)
                        {
                            check2 = true;
                            lines = res.ListLine;
                        }
                    }
                    if (check2)
                    {
                        if (points.Count >= 4)
                        {
                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                if (GeomUtil.IsEqual(points[i], points[i + 1])) continue;
                                Line l = Line.CreateBound(points[i], points[i + 1]);
                                bool check3 = true;
                                if (polygon2.IsLineInPolygon(l))
                                {
                                    bool check4 = false;
                                    for (int j = 0; j < lines.Count; j++)
                                    {
                                        LineCompareLineResult res1 = new LineCompareLineResult(lines[j], l);
                                        if (res1.Type == LineCompareLineType.SameDirectionLineOverlap) check3 = false;
                                        if (res1.Type == LineCompareLineType.SameDirectionPointOverlap)
                                        {
                                            ListLine[j] = res1.MergeLine;
                                            check4 = true; break;
                                        }
                                    }
                                    if (check4) continue;
                                    if (check3)
                                    {
                                        lines.Add(l);
                                    }
                                }
                            }
                        }
                        this.ListLine = lines;
                        this.IntersectType = PolygonComparePolygonIntersectType.Boundary; return;
                    }
                    if (check)
                    {
                        if (points.Count >= 2)
                        {
                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                if (GeomUtil.IsEqual(points[i], points[i + 1])) continue;
                                Line l = Line.CreateBound(points[i], points[i + 1]);
                                bool check3 = true;

                                if (polygon2.IsLineInPolygon(l))
                                {
                                    bool check4 = false;
                                    for (int j = 0; j < lines.Count; j++)
                                    {
                                        LineCompareLineResult res1 = new LineCompareLineResult(lines[j], l);
                                        if (res1.Type == LineCompareLineType.SameDirectionLineOverlap) check3 = false;
                                        if (res1.Type == LineCompareLineType.SameDirectionPointOverlap)
                                        {
                                            lines[j] = res1.MergeLine;
                                            check4 = true; break;
                                        }
                                    }
                                    if (check4) continue;
                                    if (check3)
                                    {
                                        lines.Add(l);
                                    }
                                }
                            }
                            if (lines.Count != 0)
                            {
                                this.ListLine = lines; this.IntersectType = PolygonComparePolygonIntersectType.Boundary; return;
                            }
                        }
                        this.ListPoint = points; this.IntersectType = PolygonComparePolygonIntersectType.Point; return;
                    }
                    this.IntersectType = PolygonComparePolygonIntersectType.NonIntersect; return;
                #endregion
                case PolygonComparePolygonPositionType.Planar:
                    check = false;
                    check2 = false;
                    List<Line> lines1 = new List<Line>(), lines2 = new List<Line>();
                    List<XYZ> points1 = new List<XYZ>(), points2 = new List<XYZ>();
                    foreach (Curve c1 in polygon1.ListCurve)
                    {
                        LineComparePolygonResult res = new LineComparePolygonResult(polygon2, CheckGeometry.ConvertLine(c1));
                        if (res.Type == LineComparePolygonType.OverlapOrIntersect || res.Type == LineComparePolygonType.Inside)
                        {
                            check2 = true;
                            foreach (Line l in res.ListLine)
                            {
                                lines1.Add(l);
                            }
                        }
                        if (res.Type == LineComparePolygonType.Outside)
                        {
                            if (polygon2.CheckXYZPointPosition(c1.GetEndPoint(0)) == PointComparePolygonResult.Boundary || polygon2.CheckXYZPointPosition(c1.GetEndPoint(0)) == PointComparePolygonResult.Node)
                            {
                                points1.Add(c1.GetEndPoint(0));
                                check = true;
                            }
                            if (polygon2.CheckXYZPointPosition(c1.GetEndPoint(1)) == PointComparePolygonResult.Boundary || polygon2.CheckXYZPointPosition(c1.GetEndPoint(1)) == PointComparePolygonResult.Node)
                            {
                                points1.Add(c1.GetEndPoint(1));
                                check = true;
                            }
                        }
                    }
                    foreach (Curve c1 in polygon2.ListCurve)
                    {
                        LineComparePolygonResult res = new LineComparePolygonResult(polygon1, CheckGeometry.ConvertLine(c1));
                        if (res.Type == LineComparePolygonType.OverlapOrIntersect || res.Type == LineComparePolygonType.Inside)
                        {
                            check2 = true;
                            foreach (Line l in res.ListLine)
                            {
                                lines2.Add(l);
                            }
                        }
                        if (res.Type == LineComparePolygonType.Outside)
                        {
                            if (polygon1.CheckXYZPointPosition(c1.GetEndPoint(0)) == PointComparePolygonResult.Boundary || polygon1.CheckXYZPointPosition(c1.GetEndPoint(0)) == PointComparePolygonResult.Node)
                            {
                                points2.Add(c1.GetEndPoint(0));
                                check = true;
                            }
                            if (polygon1.CheckXYZPointPosition(c1.GetEndPoint(1)) == PointComparePolygonResult.Boundary || polygon1.CheckXYZPointPosition(c1.GetEndPoint(1)) == PointComparePolygonResult.Node)
                            {
                                points2.Add(c1.GetEndPoint(1));
                                check = true;
                            }
                        }
                    }
                    if (check2)
                    {
                        foreach (Line l in lines2)
                        {
                            lines1.Add(l);
                        }
                        lines = new List<Line>();
                        for (int i = 0; i < lines1.Count; i++)
                        {
                            bool check3 = true;
                            for (int j = i + 1; j < lines1.Count; j++)
                            {
                                LineCompareLineResult res = new LineCompareLineResult(lines1[i], lines1[j]);
                                if (res.Type == LineCompareLineType.SameDirectionLineOverlap)
                                {
                                    check3 = false;
                                    break;
                                }
                            }
                            if (check3) lines.Add(lines1[i]);
                        }
                        this.ListLine = new List<Line>();
                        List<int> nums = new List<int>();
                        for (int i = 0; i < lines.Count; i++)
                        {
                            for (int k = 0; k < nums.Count; k++)
                            {
                                //if (i == k) goto EndLoop;
                            }
                            bool check6 = true;
                            Line temp = null;
                            for (int j = i + 1; j < lines.Count; j++)
                            {
                                LineCompareLineResult llRes = new LineCompareLineResult(lines[i], lines[j]);
                                if (llRes.Type == LineCompareLineType.SameDirectionPointOverlap)
                                {
                                    nums.Add(i); nums.Add(j);
                                    check6 = false;
                                    temp = llRes.MergeLine;
                                    break;
                                }
                            }
                            if (!check6) ListLine.Add(temp);
                            else ListLine.Add(lines[i]);
                            //EndLoop:
                            //a= 0;
                        }
                        List<Curve> cs = new List<Curve>();
                        foreach (Line l in ListLine)
                        {
                            cs.Add(l);
                        }
                        List<Polygon> pls = new List<Polygon>();
                        if (CheckGeometry.CreateListPolygon(cs, out pls))
                        {
                            ListPolygon = pls;
                            this.IntersectType = PolygonComparePolygonIntersectType.AreaOverlap; return;
                        }
                        this.IntersectType = PolygonComparePolygonIntersectType.Boundary; return;
                    }
                    if (check)
                    {
                        foreach (XYZ pnt in points2)
                        {
                            points1.Add(pnt);
                        }
                        points = new List<XYZ>();
                        for (int i = 0; i < points1.Count; i++)
                        {
                            bool check3 = true;
                            for (int j = i + 1; j < points1.Count; j++)
                            {
                                if (GeomUtil.IsEqual(points1[i], points1[j]))
                                {
                                    check3 = false; break;
                                }
                            }
                            if (check3)
                            {
                                points.Add(points1[i]);
                            }
                        }
                        this.ListPoint = points; this.IntersectType = PolygonComparePolygonIntersectType.Point; return;
                    }
                    this.IntersectType = PolygonComparePolygonIntersectType.NonIntersect; return;
            }
            throw new Exception("Code complier should never be here.");
        }
        public void GetOuterPolygon(Polygon polygonCut, out object outerPolygonOrMulti)
        {
            if (polygonCut != polygon1 && polygonCut != polygon2)
                throw new Exception("Choose polygon be cut from first two polygons!");
            if (ListPolygon[0] == polygon1 || ListPolygon[0] == polygon2)
            {
                if (ListPolygon[0] == polygonCut)
                {
                    outerPolygonOrMulti = null;
                }
                else
                {
                    outerPolygonOrMulti = new MultiPolygon(polygon1, polygon2);
                }
            }
            else
            {
                Polygon temp = polygonCut;
                foreach (Polygon pl in ListPolygon)
                {
                    temp = CheckGeometry.GetPolygonCut(temp, pl);
                }
                outerPolygonOrMulti = temp;
            }
        }
    }
    public enum PolygonComparePolygonPositionType
    {
        Planar, NonPlanar, Parallel
    }
    public enum PolygonComparePolygonIntersectType
    {
        AreaOverlap, Point, Boundary, NonIntersect
    }
}
