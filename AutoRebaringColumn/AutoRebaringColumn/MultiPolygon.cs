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
#endregion

namespace AutoRebaringColumn
{
    public class MultiPolygon
    {
        public List<XYZ> ListXYZPoint { get; private set; }
        public List<XYZ> TwoXYZPointsBoundary { get; private set; }
        public List<UV> TwoUVPointsBoundary { get; private set; }
        public XYZ CentralXYZPoint { get; private set; }
        public XYZ XVecManual { get; private set; }
        public XYZ YVecManual { get; private set; }
        public XYZ Normal { get; private set; }
        public Plane Plane { get; private set; }
        public Plane PlaneManual { get; private set; }
        public Polygon SurfacePolygon { get; private set; }
        public List<Polygon> OpeningPolygons { get; private set; }
        public MultiPolygon(PlanarFace f)
        {
            List<Polygon> pls = new List<Polygon>();
            EdgeArrayArray eAA = f.EdgeLoops;
            foreach (EdgeArray eA in eAA)
            {
                List<Curve> cs = new List<Curve>();
                foreach (Edge e in eA)
                {
                    List<XYZ> points = e.Tessellate() as List<XYZ>;
                    cs.Add(Line.CreateBound(points[0], points[1]));
                }
                pls.Add(new Polygon(cs));
                if (eAA.Size == 1)
                {
                    SurfacePolygon = pls[0];
                    OpeningPolygons = new List<Polygon>();
                    goto GetParameters;
                }
            }
            for (int i = 0; i < pls.Count; i++)
            {
                Plane plane = pls[i].Plane;
                for (int j = i + 1; j < pls.Count; j++)
                {
                    Polygon tempPoly = CheckGeometry.GetProjectPolygon(plane, pls[j]);
                    PolygonComparePolygonResult res = new PolygonComparePolygonResult(pls[i], tempPoly);
                    if (res.IntersectType == PolygonComparePolygonIntersectType.AreaOverlap)
                    {
                        if (res.ListPolygon[0] == pls[i])
                        {
                            SurfacePolygon = pls[j];
                            goto FinishLoops;
                        }
                        if (res.ListPolygon[0] == pls[j])
                        {
                            SurfacePolygon = pls[i];
                            goto FinishLoops;
                        }
                        else throw new Exception("Face must contain polygons inside polygon!");
                    }
                }
            }
            FinishLoops:
            if (SurfacePolygon == null) throw new Exception("Error when retrieve surface polygon!");
            Plane = SurfacePolygon.Plane;
            OpeningPolygons = new List<Polygon>();
            foreach (Polygon pl in pls)
            {
                if (pl == SurfacePolygon) continue;
                Polygon tempPoly = CheckGeometry.GetProjectPolygon(Plane, pl);
                OpeningPolygons.Add(tempPoly);
            }

            GetParameters:
            GetParameters();
        }
        public MultiPolygon(Polygon surPolygon, List<Polygon> openPolygons)
        {
            SurfacePolygon = surPolygon;
            OpeningPolygons = openPolygons;
            
            GetParameters();
        }
        public MultiPolygon(Polygon surPolygon, Polygon openPolygon)
        {
            SurfacePolygon = surPolygon;
            OpeningPolygons = new List<Polygon> { openPolygon };

            GetParameters();
        }
        private void GetParameters()
        {
            ListXYZPoint = SurfacePolygon.ListXYZPoint;
            Normal = SurfacePolygon.Normal;
            CentralXYZPoint = SurfacePolygon.CentralXYZPoint;
        }
        public void SetManualDirection(XYZ vec, bool isXVector = true)
        {
            SurfacePolygon.SetManualDirection(vec, isXVector);
            this.XVecManual = SurfacePolygon.XVecManual;
            this.YVecManual = SurfacePolygon.YVecManual;
            this.PlaneManual = SurfacePolygon.PlaneManual;
        }
        public void SetTwoPointsBoundary(XYZ vec, bool isXVector = true)
        {
            SetManualDirection(vec, isXVector);
            SurfacePolygon.SetTwoPointsBoundary(vec, isXVector);
            this.TwoUVPointsBoundary = SurfacePolygon.TwoUVPointsBoundary;
            this.TwoXYZPointsBoundary = SurfacePolygon.TwoXYZPointsBoundary;
        }
    }
    public class PolygonCompareMultiPolygonResult
    {
        public PolygonCompareMultiPolygonPositionType PositionType { get; private set; }
        public PolygonCompareMultiPolygonIntersectType IntersectType { get; private set; }
        public List<Line> ListLine { get; private set; }
        public List<XYZ> ListPoint { get; private set; }
        public List<Polygon> ListPolygon { get; private set; }
        private Polygon polygon;
        private MultiPolygon multiPolygon;
        public MultiPolygon MultiPolygon;
        public PolygonCompareMultiPolygonResult(Polygon pl, MultiPolygon mpl)
        {
            this.polygon = pl; this.multiPolygon = mpl; GetPositionType(); GetIntersectTypeAndOtherParameter();
        }
        private void GetPositionType()
        {
            if (GeomUtil.IsSameOrOppositeDirection(polygon.Normal, multiPolygon.Normal))
            {
                if (multiPolygon.SurfacePolygon.CheckXYZPointPosition(polygon.ListXYZPoint[0]) != PointComparePolygonResult.NonPlanar)
                {
                    PositionType = PolygonCompareMultiPolygonPositionType.Planar;
                    return;
                }
                PositionType = PolygonCompareMultiPolygonPositionType.Parallel; return;
            }
            PositionType = PolygonCompareMultiPolygonPositionType.NonPlarnar; return;
        }
        private void GetIntersectTypeAndOtherParameter()
        {
            switch (PositionType)
            {
                case PolygonCompareMultiPolygonPositionType.Parallel: this.IntersectType = PolygonCompareMultiPolygonIntersectType.NonIntersect; return;
                case PolygonCompareMultiPolygonPositionType.NonPlarnar: throw new Exception("Code for this case hasn't finished yet!");
                case PolygonCompareMultiPolygonPositionType.Planar:
                    Polygon surPL = multiPolygon.SurfacePolygon;
                    List<Polygon> openPLs = multiPolygon.OpeningPolygons;
                    PolygonComparePolygonResult res = new PolygonComparePolygonResult(polygon, surPL);
                    switch (res.IntersectType)
                    {
                        case PolygonComparePolygonIntersectType.NonIntersect: this.IntersectType = PolygonCompareMultiPolygonIntersectType.NonIntersect; return;
                        case PolygonComparePolygonIntersectType.Boundary: this.IntersectType = PolygonCompareMultiPolygonIntersectType.Boundary; ListLine = res.ListLine; return;
                        case PolygonComparePolygonIntersectType.Point: IntersectType = PolygonCompareMultiPolygonIntersectType.Point; ListPoint = res.ListPoint; return;
                        case PolygonComparePolygonIntersectType.AreaOverlap:
                            ListPolygon = new List<Polygon>();
                            this.MultiPolygon = null;
                            foreach (Polygon pl in res.ListPolygon)
                            {
                                Polygon temp = pl;
                                foreach (Polygon openPl in multiPolygon.OpeningPolygons)
                                {
                                    PolygonComparePolygonResult ppRes = new PolygonComparePolygonResult(temp, openPl);
                                    if (ppRes.IntersectType == PolygonComparePolygonIntersectType.AreaOverlap)
                                    {
                                        object polyorMultiPolygonCut = null;
                                        ppRes.GetOuterPolygon(temp, out polyorMultiPolygonCut);
                                        if (polyorMultiPolygonCut == null)
                                            goto Here;
                                        if (polyorMultiPolygonCut is MultiPolygon)
                                        {
                                            this.MultiPolygon = polyorMultiPolygonCut as MultiPolygon;
                                            return;
                                        }
                                        temp = polyorMultiPolygonCut as Polygon;
                                    }
                                }
                                ListPolygon.Add(temp);
                                Here: continue;
                            }
                            break;
                    }
                    break;
            }
        }
    }
    public enum PolygonCompareMultiPolygonPositionType
    {
        Planar, NonPlarnar, Parallel
    }
    public enum PolygonCompareMultiPolygonIntersectType
    {
        AreaOverlap, Point, Boundary, NonIntersect
    }
}
