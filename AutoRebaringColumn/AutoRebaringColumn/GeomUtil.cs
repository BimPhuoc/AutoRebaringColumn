#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
#endregion

namespace AutoRebaringWall
{
    class GeomUtil
    {
        const string revit = "Revit";
        const double Precision = 0.00001;    //precision when judge whether two doubles are equal
        const double FEET_TO_METERS = 0.3048;
        const double FEET_TO_CENTIMETERS = FEET_TO_METERS * 100;
        const double FEET_TO_MILIMETERS = FEET_TO_METERS * 1000;
        public static double feet2Meter(double feet)
        {
            return feet * FEET_TO_METERS;
        }
        public static double feet2Milimeter(double feet)
        {
            return feet * FEET_TO_MILIMETERS;
        }
        public static double meter2Feet(double meter)
        {
            return meter / FEET_TO_METERS;
        }
        public static double milimeter2Feet(double milimeter)
        {
            return milimeter / FEET_TO_MILIMETERS;
        }
        public static double radian2Degree(double rad)
        {
            return rad * 180 / Math.PI;
        }
        public static double degree2Radian(double deg)
        {
            return deg * Math.PI / 180;
        }
        public static bool IsEqual(double d1, double d2)
        {
            //get the absolute value;
            double diff = Math.Abs(d1 - d2);
            return diff < Precision;
        }
        public static bool IsEqual(Autodesk.Revit.DB.XYZ first, Autodesk.Revit.DB.XYZ second)
        {
            bool flag = true;
            flag = flag && IsEqual(first.X, second.X);
            flag = flag && IsEqual(first.Y, second.Y);
            flag = flag && IsEqual(first.Z, second.Z);
            return flag;
        }
        public static bool IsEqual(Autodesk.Revit.DB.UV first, Autodesk.Revit.DB.UV second)
        {
            bool flag = true;
            flag = flag && IsEqual(first.U, second.U);
            flag = flag && IsEqual(first.V, second.V);
            return flag;
        }
        public static bool IsEqual(Curve first, Curve second)
        {
            if (IsEqual(first.GetEndPoint(0), second.GetEndPoint(0)))
            {
                return IsEqual(first.GetEndPoint(1), second.GetEndPoint(1));
            }
            if (IsEqual(first.GetEndPoint(1), second.GetEndPoint(0)))
            {
                return IsEqual(first.GetEndPoint(0), second.GetEndPoint(1));
            }
            return false;
        }
        public static bool IsSmaller(XYZ first, XYZ second)
        {
            if (IsEqual(first, second)) return false;
            if (IsEqual(first.Z, second.Z))
            {
                if (IsEqual(first.Y, second.Y))
                {
                    return (first.X < second.X);
                }
                return (first.Y < second.Y);
            }
            return (first.Z < second.Z);
        }
        public static bool IsSmaller(double x, double y)
        {
            if (IsEqual(x, y)) return false;
            return x < y;
        }

        public static bool IsBigger(XYZ first, XYZ second)
        {
            if (IsEqual(first, second)) return false;
            if (IsEqual(first.Z, second.Z))
            {
                if (IsEqual(first.Y, second.Y))
                {
                    return (first.X > second.X);
                }
                return (first.Y > second.Y);
            }
            return (first.Z > second.Z);
        }
        public static bool IsBigger(double first, double second)
        {
            if (IsEqual(first, second)) return false;
            return first > second;
        }

        public static bool IsVertical(Face face, Line line, Transform faceTrans, Transform lineTrans)
        {
            //get points which the face contains
            List<XYZ> points = face.Triangulate().Vertices as List<XYZ>;
            if (3 > points.Count)    // face's point number should be above 2
            {
                return false;
            }

            // get three points from the face points
            Autodesk.Revit.DB.XYZ first = points[0];
            Autodesk.Revit.DB.XYZ second = points[1];
            Autodesk.Revit.DB.XYZ third = points[2];

            // get start and end point of line
            Autodesk.Revit.DB.XYZ lineStart = line.GetEndPoint(0);
            Autodesk.Revit.DB.XYZ lineEnd = line.GetEndPoint(1);

            // transForm the three points if necessary
            if (null != faceTrans)
            {
                first = TransformPoint(first, faceTrans);
                second = TransformPoint(second, faceTrans);
                third = TransformPoint(third, faceTrans);
            }

            // transform the start and end points if necessary
            if (null != lineTrans)
            {
                lineStart = TransformPoint(lineStart, lineTrans);
                lineEnd = TransformPoint(lineEnd, lineTrans);
            }

            // form two vectors from the face and a vector stand for the line
            // Use SubXYZ() method to get the vectors
            Autodesk.Revit.DB.XYZ vector1 = SubXYZ(first, second);    // first vector of face
            Autodesk.Revit.DB.XYZ vector2 = SubXYZ(first, third);     // second vector of face
            Autodesk.Revit.DB.XYZ vector3 = SubXYZ(lineStart, lineEnd);   // line vector

            // get two dot products of the face vectors and line vector
            double result1 = DotMatrix(vector1, vector3);
            double result2 = DotMatrix(vector2, vector3);

            // if two dot products are all zero, the line is perpendicular to the face
            return (IsEqual(result1, 0) && IsEqual(result2, 0));
        }
        public static bool IsSameDirection(Autodesk.Revit.DB.XYZ firstVec, Autodesk.Revit.DB.XYZ secondVec)
        {
            Autodesk.Revit.DB.XYZ first = UnitVector(firstVec);
            Autodesk.Revit.DB.XYZ second = UnitVector(secondVec);
            double dot = DotMatrix(first, second);
            return (IsEqual(dot, 1));
        }
        public static bool IsSameDirection(UV firstVec, UV secondVec)
        {
            Autodesk.Revit.DB.UV first = UnitVector(firstVec);
            Autodesk.Revit.DB.UV second = UnitVector(secondVec);
            double dot = DotMatrix(first, second);
            return (IsEqual(dot, 1));
        }
        public static bool IsSameOrOppositeDirection(XYZ firstVec, XYZ secondVec)
        {
            Autodesk.Revit.DB.XYZ first = UnitVector(firstVec);
            Autodesk.Revit.DB.XYZ second = UnitVector(secondVec);

            // if the dot product of two unit vectors is equal to 1, return true
            double dot = DotMatrix(first, second);
            return (IsEqual(dot, 1) || IsEqual(dot, -1));
        }
        public static bool IsSameOrOppositeDirection(UV firstVec, UV secondVec)
        {
            Autodesk.Revit.DB.UV first = UnitVector(firstVec);
            Autodesk.Revit.DB.UV second = UnitVector(secondVec);

            // if the dot product of two unit vectors is equal to 1, return true
            double dot = DotMatrix(first, second);
            return (IsEqual(dot, 1) || IsEqual(dot, -1));
        }
        public static bool IsOppositeDirection(Autodesk.Revit.DB.XYZ firstVec, Autodesk.Revit.DB.XYZ secondVec)
        {
            // get the unit vector for two vectors
            Autodesk.Revit.DB.XYZ first = UnitVector(firstVec);
            Autodesk.Revit.DB.XYZ second = UnitVector(secondVec);
            // if the dot product of two unit vectors is equal to -1, return true
            double dot = DotMatrix(first, second);
            return (IsEqual(dot, -1));
        }
        public static bool IsOppositeDirection(Autodesk.Revit.DB.UV firstVec, Autodesk.Revit.DB.UV secondVec)
        {
            // get the unit vector for two vectors
            Autodesk.Revit.DB.UV first = UnitVector(firstVec);
            Autodesk.Revit.DB.UV second = UnitVector(secondVec);

            // if the dot product of two unit vectors is equal to -1, return true
            double dot = DotMatrix(first, second);
            return (IsEqual(dot, -1));
        }
        public static Autodesk.Revit.DB.XYZ CrossMatrix(Autodesk.Revit.DB.XYZ p1, Autodesk.Revit.DB.XYZ p2)
        {
            //get the coordinate of the XYZ
            double u1 = p1.X;
            double u2 = p1.Y;
            double u3 = p1.Z;

            double v1 = p2.X;
            double v2 = p2.Y;
            double v3 = p2.Z;

            double x = v3 * u2 - v2 * u3;
            double y = v1 * u3 - v3 * u1;
            double z = v2 * u1 - v1 * u2;

            return new Autodesk.Revit.DB.XYZ(x, y, z);
        }
        public static Autodesk.Revit.DB.XYZ UnitVector(Autodesk.Revit.DB.XYZ vector)
        {
            // calculate the distance from grid origin to the XYZ
            double length = GetLength(vector);

            // changed the vector into the unit length
            double x = vector.X / length;
            double y = vector.Y / length;
            double z = vector.Z / length;
            return new Autodesk.Revit.DB.XYZ(x, y, z);
        }
        public static Autodesk.Revit.DB.UV UnitVector(Autodesk.Revit.DB.UV vector)
        {
            // calculate the distance from grid origin to the XYZ
            double length = GetLength(vector);

            // changed the vector into the unit length
            double x = vector.U / length;
            double y = vector.V / length;
            return new Autodesk.Revit.DB.UV(x, y);
        }
        public static double GetLength(Autodesk.Revit.DB.XYZ vector)
        {
            double x = vector.X;
            double y = vector.Y;
            double z = vector.Z;
            return Math.Sqrt(x * x + y * y + z * z);
        }
        public static double GetLength(Autodesk.Revit.DB.XYZ vector, bool checkPositive)
        {
            double len = GetLength(vector);
            if (!checkPositive) return len;
            return IsSmaller(-vector, vector) ? len : -len;
        }

        public static double GetLength(Autodesk.Revit.DB.UV vector)
        {
            double x = vector.U;
            double y = vector.V;
            return Math.Sqrt(x * x + y * y);
        }
        public static double GetLength(Line line)
        {
            XYZ first = line.GetEndPoint(0);
            XYZ second = line.GetEndPoint(1);
            XYZ vec = SubXYZ(first, second);
            return GetLength(vec);
        }
        public static double GetLength(XYZ p1, XYZ p2)
        {
            return GetLength(SubXYZ(p1, p2));
        }
        public static XYZ GetMiddlePoint(XYZ first, XYZ second)
        {
            return (first + second) / 2;
        }
        public static Autodesk.Revit.DB.XYZ SubXYZ(Autodesk.Revit.DB.XYZ p1, Autodesk.Revit.DB.XYZ p2)
        {
            double x = p1.X - p2.X;
            double y = p1.Y - p2.Y;
            double z = p1.Z - p2.Z;

            return new Autodesk.Revit.DB.XYZ(x, y, z);
        }
        public static Autodesk.Revit.DB.UV SubXYZ(Autodesk.Revit.DB.UV p1, Autodesk.Revit.DB.UV p2)
        {
            double x = p1.U - p2.U;
            double y = p1.V - p2.V;

            return new Autodesk.Revit.DB.UV(x, y);
        }
        public static Autodesk.Revit.DB.XYZ AddXYZ(Autodesk.Revit.DB.XYZ p1, Autodesk.Revit.DB.XYZ p2)
        {
            double x = p1.X + p2.X;
            double y = p1.Y + p2.Y;
            double z = p1.Z + p2.Z;

            return new Autodesk.Revit.DB.XYZ(x, y, z);
        }
        public static Autodesk.Revit.DB.UV AddXYZ(Autodesk.Revit.DB.UV p1, Autodesk.Revit.DB.UV p2)
        {
            double x = p1.U + p2.U;
            double y = p1.V + p2.V;

            return new Autodesk.Revit.DB.UV(x, y);
        }
        public static Autodesk.Revit.DB.XYZ MultiplyVector(Autodesk.Revit.DB.XYZ vector, double rate)
        {
            double x = vector.X * rate;
            double y = vector.Y * rate;
            double z = vector.Z * rate;

            return new Autodesk.Revit.DB.XYZ(x, y, z);
        }
        public static Autodesk.Revit.DB.XYZ TransformPoint(Autodesk.Revit.DB.XYZ point, Transform transform)
        {
            //get the coordinate value in X, Y, Z axis
            double x = point.X;
            double y = point.Y;
            double z = point.Z;

            //transform basis of the old coordinate system in the new coordinate system
            Autodesk.Revit.DB.XYZ b0 = transform.get_Basis(0);
            Autodesk.Revit.DB.XYZ b1 = transform.get_Basis(1);
            Autodesk.Revit.DB.XYZ b2 = transform.get_Basis(2);
            Autodesk.Revit.DB.XYZ origin = transform.Origin;

            //transform the origin of the old coordinate system in the new coordinate system
            double xTemp = x * b0.X + y * b1.X + z * b2.X + origin.X;
            double yTemp = x * b0.Y + y * b1.Y + z * b2.Y + origin.Y;
            double zTemp = x * b0.Z + y * b1.Z + z * b2.Z + origin.Z;

            return new Autodesk.Revit.DB.XYZ(xTemp, yTemp, zTemp);
        }
        public static Autodesk.Revit.DB.Curve TransformCurve(Autodesk.Revit.DB.Curve curve, Transform transform)
        {
            return Line.CreateBound(TransformPoint(curve.GetEndPoint(0), transform), TransformPoint(curve.GetEndPoint(1), transform));
        }
        public static Autodesk.Revit.DB.XYZ OffsetPoint(Autodesk.Revit.DB.XYZ point, Autodesk.Revit.DB.XYZ direction, double offset)
        {
            Autodesk.Revit.DB.XYZ directUnit = UnitVector(direction);
            Autodesk.Revit.DB.XYZ offsetVect = MultiplyVector(directUnit, offset);
            return AddXYZ(point, offsetVect);
        }
        public static Autodesk.Revit.DB.Curve OffsetCurve(Autodesk.Revit.DB.Curve c, Autodesk.Revit.DB.XYZ direction, double offset)
        {
            c = Line.CreateBound(OffsetPoint(c.GetEndPoint(0), direction, offset), OffsetPoint(c.GetEndPoint(1), direction, offset));
            return c;
        }

        public static List<Curve> OffsetListCurve(List<Curve> cs, Autodesk.Revit.DB.XYZ direction, double offset)
        {
            for (int i = 0; i < cs.Count; i++)
            {
                cs[i] = OffsetCurve(cs[i], direction, offset);
            }
            return cs;
        }
        public static Polygon OffsetPolygon(Polygon pl, Autodesk.Revit.DB.XYZ direction, double offset)
        {
            List<Curve> cs = new List<Curve>();
            foreach (Curve c in pl.ListCurve)
            {
                Curve c1 = OffsetCurve(c, direction, offset);
                cs.Add(c1);
            }
            return new Polygon(cs);
        }
        public static RebarHookOrientation GetHookOrient(Autodesk.Revit.DB.XYZ curveVec, Autodesk.Revit.DB.XYZ normal, Autodesk.Revit.DB.XYZ hookVec)
        {
            Autodesk.Revit.DB.XYZ tempVec = normal;

            for (int i = 0; i < 4; i++)
            {
                tempVec = GeomUtil.CrossMatrix(tempVec, curveVec);
                if (GeomUtil.IsSameDirection(tempVec, hookVec))
                {
                    if (i == 0)
                    {
                        return RebarHookOrientation.Right;
                    }
                    else if (i == 2)
                    {
                        return RebarHookOrientation.Left;
                    }
                }
            }

            throw new Exception("Can't find the hook orient according to hook direction.");
        }
        public static bool IsInRightDir(Autodesk.Revit.DB.XYZ normal)
        {
            double eps = 1.0e-8;
            if (Math.Abs(normal.X) <= eps)
            {
                if (normal.Y > 0) return false;
                else return true;
            }
            if (normal.X > 0) return true;
            if (normal.X < 0) return false;
            return true;
        }
        public static double DotMatrix(Autodesk.Revit.DB.XYZ p1, Autodesk.Revit.DB.XYZ p2)
        {
            //get the coordinate of the Autodesk.Revit.DB.XYZ 
            double v1 = p1.X;
            double v2 = p1.Y;
            double v3 = p1.Z;

            double u1 = p2.X;
            double u2 = p2.Y;
            double u3 = p2.Z;

            return v1 * u1 + v2 * u2 + v3 * u3;
        }
        public static double DotMatrix(Autodesk.Revit.DB.UV p1, Autodesk.Revit.DB.UV p2)
        {
            //get the coordinate of the Autodesk.Revit.DB.XYZ 
            double v1 = p1.U;
            double v2 = p1.V;


            double u1 = p2.U;
            double u2 = p2.V;


            return v1 * u1 + v2 * u2;
        }

        public static int RoundUp(double d)
        {
            return Math.Round(d, 0) < d ? (int)(Math.Round(d, 0) + 1) : (int)(Math.Round(d, 0));
        }
        public static int RoundDown(double d)
        {
            return Math.Round(d, 0) < d ? (int)(Math.Round(d, 0)) : (int)(Math.Round(d, 0) - 1);
        }

        public static double GetRadianAngle(XYZ vec1, XYZ vec2)
        {
            return Math.Acos(DotMatrix(UnitVector(vec1), UnitVector(vec2)));
        }
        public static double GetDegreeAngle(XYZ vec1, XYZ vec2)
        {
            return radian2Degree(GetRadianAngle(vec1, vec2));
        }
    }
    public class ZYXComparer : IComparer<XYZ>
    {
        int IComparer<XYZ>.Compare(XYZ first, XYZ second)
        {
            // first compare z coordinate, then y coordiante, at last x coordinate
            if (GeomUtil.IsEqual(first.Z, second.Z))
            {
                if (GeomUtil.IsEqual(first.Y, second.Y))
                {
                    if (GeomUtil.IsEqual(first.X, second.X))
                    {
                        return 0; // Equal
                    }
                    return (first.X > second.X) ? 1 : -1;
                }
                return (first.Y > second.Y) ? 1 : -1;
            }
            return (first.Z > second.Z) ? 1 : -1;
        }
    }
    public class YXComparer : IComparer<XYZ>
    {
        int IComparer<XYZ>.Compare(XYZ first, XYZ second)
        {
            if (GeomUtil.IsEqual(first.Y, second.Y))
            {
                if (GeomUtil.IsEqual(first.X, second.X))
                {
                    return 0; // Equal
                }
                return (first.X > second.X) ? 1 : -1;
            }
            return (first.Y > second.Y) ? 1 : -1;
        }
    }
    public class XYComparer : IComparer<XYZ>
    {
        int IComparer<XYZ>.Compare(XYZ first, XYZ second)
        {
            if (GeomUtil.IsEqual(first.X, second.X))
            {
                if (GeomUtil.IsEqual(first.Y, second.Y))
                {
                    return 0; // Equal
                }
                return (first.Y > second.Y) ? 1 : -1;
            }
            return (first.X > second.X) ? 1 : -1;
        }
    }
    public class ElementLocationComparer : IComparer<Element>
    {
        private Plane Plane;
        public ElementLocationComparer()
        {

        }
        public ElementLocationComparer(Plane pl)
        {
            Plane = pl;
        }
        int IComparer<Element>.Compare(Element x, Element y)
        {
            XYZ loc1 = null;
            XYZ loc2 = null;
            if (x is Wall)
            {
                WallGeometryInfo wgi = new WallGeometryInfo(x as Wall); loc1 = wgi.TopPolygon.CentralXYZPoint;
            }
            if (x is FamilyInstance)
            {
                ColumnGeometryInfo cgi = new ColumnGeometryInfo(x as FamilyInstance); loc1 = cgi.TopPolygon.CentralXYZPoint;
            }
            if (y is Wall)
            {
                WallGeometryInfo wgi = new WallGeometryInfo(y as Wall); loc2 = wgi.TopPolygon.CentralXYZPoint;
            }
            if (y is FamilyInstance)
            {
                ColumnGeometryInfo cgi = new ColumnGeometryInfo(y as FamilyInstance); loc2 = cgi.TopPolygon.CentralXYZPoint;
            }
            loc1 = new XYZ(loc1.X, loc1.Y, 0); loc2 = new XYZ(loc2.X, loc2.Y, 0);
            if (Plane != null)
            {
                loc1 = CheckGeometry.GetProjectPoint(Plane, loc1); loc2 = CheckGeometry.GetProjectPoint(Plane, loc2);
            }

            if (GeomUtil.IsEqual(loc1.X, loc2.X))
            {
                if (GeomUtil.IsEqual(loc1.Y, loc2.Y))
                {
                    return 0; // Equal
                }
                return (loc1.Y > loc2.Y) ? 1 : -1;
            }
            return (loc1.X > loc2.X) ? 1 : -1;
        }
    }
}
