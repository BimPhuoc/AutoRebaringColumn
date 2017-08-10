#region Namespaces
using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using AutoRebaringWall;
using DataExcel;using Export;

#endregion
namespace Phuoc
{
    public class StandardBarWall
    {
        UIApplication uiapp;
        UIDocument uidoc;
        Application app;
        Document doc;
        Selection sel;
        View activeView;
        const double f2m = 0.3048;
        const double f2mm = f2m * 1000;
        const double mm2f = 1 / f2mm;
        public enum draw
        {
            straight,
            hook,
            zigzag,
            nodraw,
            vuottang
        };
        public StandardBarWall(ExternalCommandData commandData, ref string message, ElementSet element, string path)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;
            sel = uidoc.Selection;
            activeView = uidoc.ActiveView;
            ////filter walls frome selection , Sort List Wall folllow Vec_Z
            List<Wall> wall = sortWallByLevel(sel);
            #region "INPUT EXCEL"
            ImportExcel excel = new ImportExcel(wall.Count, path);
            DataProcess data = new DataProcess(excel, commandData, wall);
            #endregion
            #region "PHỐI THÉP"
            List<floorOnWall> san = new List<floorOnWall>();
            for (int i = 0; i < wall.Count; i++) san.Add(new floorOnWall(wall[i], doc, sel));
            CombineRebar combineRebar_bien = new CombineRebar(excel, data.hVach, san, excel.bt, excel.bd, data.D_bien, excel.dtc1_bien, excel.dtc2_bien);
            CombineRebar combineRebar_giua = new CombineRebar(excel, data.hVach, san, excel.bt, excel.bd, data.D_giua, excel.dtc1_giua, excel.dtc2_giua);
            #endregion
            #region "export PHỐI THÉP"
            //xuất kết quả sau khi phối thép
            List<double> L1_bien = combineRebar_bien.L1;
            List<double> L2_bien = combineRebar_bien.L2;
            List<double> L0_bien = combineRebar_bien.L0;
            List<string> Comment_bien = combineRebar_bien.comment;
            List<bool> vt1_bien = combineRebar_bien.vuotTang1;
            List<bool> vt2_bien = combineRebar_bien.vuotTang2;

            List<double> L1_giua = combineRebar_giua.L1;
            List<double> L2_giua = combineRebar_giua.L2;
            List<double> L0_giua = combineRebar_giua.L0;
            List<string> Comment_giua = combineRebar_giua.comment;
            List<bool> vt1_giua = combineRebar_giua.vuotTang1;
            List<bool> vt2_giua = combineRebar_giua.vuotTang2;
            #endregion
            #region "điểm offset các điểm bắt vẽ thép"
            double[] z1_bien = z_offset(L1_bien, vt1_bien, data.D_bien, data.hVach, excel.dtc1_bien, excel.n1_);
            double[] z2_bien = z_offset(L2_bien, vt2_bien, data.D_bien, data.hVach, excel.dtc2_bien, excel.n1_);
            double[] z1_giua = z_offset(L1_giua, vt1_giua, data.D_giua, data.hVach, excel.dtc1_giua, excel.n1_);
            double[] z2_giua = z_offset(L2_giua, vt2_giua, data.D_giua, data.hVach, excel.dtc2_giua, excel.n1_);
            #endregion
            #region"VẼ THÉP"
            for (int i = 0; i < wall.Count; i++)
            {
                #region "input"
                List<XYZ> bp0 = data.bp0; List<XYZ> mp01 = data.mp01; List<XYZ> mp12 = data.mp12;
                List<XYZ> mp02 = data.mp02; List<XYZ> bp1 = data.bp1; List<XYZ> mp11 = data.mp11;
                List<XYZ> vecX = data.vecX; List<XYZ> vecY = data.vecY;
                List<int> number_A = excel.number_A; List<int> number_B = excel.number_B; List<int> number_C = excel.number_C;
                List<double> A = data.A; List<double> B = data.B; List<double> C = data.C;
                List<double> D_bien = data.D_bien; List<double> D_giua = data.D_giua;
                List<RebarBarType> type_D_bien = data.type_D_bien; List<RebarBarType> type_D_giua = data.type_D_giua;
                #endregion
                #region "vẽ đặc biệt"
                #region "Rebar Point of Wall1"
                //vách[i]
                RebarPoint w1_c1 = new RebarPoint(true, bp0[i], vecX[i], number_A[i], A[i], D_bien[i], type_D_bien[i], L1_bien[i], L2_bien[i], z1_bien[i], z2_bien[i], false);
                RebarPoint w1_c4 = new RebarPoint(false, mp01[i], -vecY[i], number_C[i], C[i], D_bien[i], type_D_bien[i], L1_bien[i], L2_bien[i], z1_bien[i], z2_bien[i], w1_c1.sole_cuoi);
                RebarPoint w1_c5 = new RebarPoint(true, mp12[i], -vecX[i], number_A[i], A[i], D_bien[i], type_D_bien[i], L1_bien[i], L2_bien[i], z1_bien[i], z2_bien[i], w1_c4.sole_cuoi);
                RebarPoint w1_c2 = new RebarPoint(false, bp0[i], -vecY[i], number_C[i], C[i], D_bien[i], type_D_bien[i], L1_bien[i], L2_bien[i], z1_bien[i], z2_bien[i], w1_c5.sole_cuoi);
                RebarPoint w1_c3 = new RebarPoint(false, mp01[i], vecX[i], number_B[i], B[i], D_giua[i], type_D_giua[i], L1_giua[i], L2_giua[i], z1_giua[i], z2_giua[i], w1_c1.sole_cuoi);

                RebarPoint w1_c10 = new RebarPoint(true, mp02[i], vecX[i], number_A[i], A[i], D_bien[i], type_D_bien[i], L1_bien[i], L2_bien[i], z1_bien[i], z2_bien[i], w1_c3.sole_cuoi);
                RebarPoint w1_c7 = new RebarPoint(false, bp1[i], vecY[i], number_C[i], C[i], D_bien[i], type_D_bien[i], L1_bien[i], L2_bien[i], z1_bien[i], z2_bien[i], w1_c10.sole_cuoi);
                RebarPoint w1_c6 = new RebarPoint(true, bp1[i], -vecX[i], number_A[i], A[i], D_bien[i], type_D_bien[i], L1_bien[i], L2_bien[i], z1_bien[i], z2_bien[i], w1_c7.sole_cuoi);
                RebarPoint w1_c9 = new RebarPoint(false, mp11[i], vecY[i], number_C[i], C[i], D_bien[i], type_D_bien[i], L1_bien[i], L2_bien[i], z1_bien[i], z2_bien[i], w1_c6.sole_cuoi);
                RebarPoint w1_c8 = new RebarPoint(false, mp11[i], -vecX[i], number_B[i], B[i], D_giua[i], type_D_giua[i], L1_giua[i], L2_giua[i], z1_giua[i], z2_giua[i], w1_c6.sole_cuoi);

                RebarPoint w1_l1 = w1_c1.Add(w1_c3).Add(w1_c10);
                RebarPoint w1_l2 = w1_c6.Add(w1_c8).Add(w1_c5);

                #endregion
                //vách[i+1]
                int j2 = i + 1;
                if (i <wall.Count-1)
                {
                    #region "Rebar Point"
                    RebarPoint w2_c1 = new RebarPoint(true, bp0[j2], vecX[j2], number_A[j2], A[j2], D_bien[j2], type_D_bien[j2], L1_bien[j2], L2_bien[j2], z1_bien[j2], z2_bien[j2], false);
                    RebarPoint w2_c4 = new RebarPoint(false, mp01[j2], -vecY[j2], number_C[j2], C[j2], D_bien[j2], type_D_bien[j2], L1_bien[j2], L2_bien[j2], z1_bien[j2], z2_bien[j2], w2_c1.sole_cuoi);
                    RebarPoint w2_c5 = new RebarPoint(true, mp12[j2], -vecX[j2], number_A[j2], A[j2], D_bien[j2], type_D_bien[j2], L1_bien[j2], L2_bien[j2], z1_bien[j2], z2_bien[j2], w2_c4.sole_cuoi);
                    RebarPoint w2_c2 = new RebarPoint(false, bp0[j2], -vecY[j2], number_C[j2], C[j2], D_bien[j2], type_D_bien[j2], L1_bien[j2], L2_bien[j2], z1_bien[j2], z2_bien[j2], w2_c5.sole_cuoi);
                    RebarPoint w2_c3 = new RebarPoint(false, mp01[j2], vecX[j2], number_B[j2], B[j2], D_giua[j2], type_D_giua[j2], L1_giua[j2], L2_giua[j2], z1_giua[j2], z2_giua[j2], w2_c1.sole_cuoi);

                    RebarPoint w2_c10 = new RebarPoint(true, mp02[j2], vecX[j2], number_A[j2], A[j2], D_bien[j2], type_D_bien[j2], L1_bien[j2], L2_bien[j2], z1_bien[j2], z2_bien[j2], w2_c3.sole_cuoi);
                    RebarPoint w2_c7 = new RebarPoint(false, bp1[j2], vecY[j2], number_C[j2], C[j2], D_bien[j2], type_D_bien[j2], L1_bien[j2], L2_bien[j2], z1_bien[j2], z2_bien[j2], w2_c10.sole_cuoi);
                    RebarPoint w2_c6 = new RebarPoint(true, bp1[j2], -vecX[j2], number_A[j2], A[j2], D_bien[j2], type_D_bien[j2], L1_bien[j2], L2_bien[j2], z1_bien[j2], z2_bien[j2], w2_c7.sole_cuoi);
                    RebarPoint w2_c9 = new RebarPoint(false, mp11[j2], vecY[j2], number_C[j2], C[j2], D_bien[j2], type_D_bien[j2], L1_bien[j2], L2_bien[j2], z1_bien[j2], z2_bien[j2], w2_c6.sole_cuoi);
                    RebarPoint w2_c8 = new RebarPoint(false, mp11[j2], -vecX[j2], number_B[j2], B[j2], D_giua[j2], type_D_giua[j2], L1_giua[j2], L2_giua[j2], z1_giua[j2], z2_giua[j2], w2_c6.sole_cuoi);

                    RebarPoint w2_l1 = w2_c1.Add(w2_c3).Add(w2_c10);
                    RebarPoint w2_l2 = w2_c6.Add(w2_c8).Add(w2_c5);
                    #endregion
                    #region "Xác định cách vẽ ứng với Rebar Point"
                    DentifyRebar(w1_l1, w2_l1, wall[i + 1], data.a[i+1], excel);
                    //DentifyRebar(wall1_l2, wall2_l2, wall[i + 1], data.a[i], excel);
                    //DentifyRebar(wall1_c2, wall2_c2, wall[i + 1], data.a[i], excel);
                    //DentifyRebar(wall1_c4, wall2_c4, wall[i + 1], data.a[i], excel);
                    //DentifyRebar(wall1_c9, wall2_c9, wall[i + 1], data.a[i], excel);
                    //DentifyRebar(wall1_c7, wall2_c7, wall[i + 1], data.a[i], excel);
                    #endregion
                }
                else
                {
                    #region "Dentify Rebar"
                    DentifyRebar(w1_l1, draw.hook);
                    DentifyRebar(w1_l2, draw.hook);
                    DentifyRebar(w1_c2, draw.hook);
                    DentifyRebar(w1_c4, draw.hook);
                    DentifyRebar(w1_c9, draw.hook);
                    DentifyRebar(w1_c7, draw.hook);
                    #endregion
                }
                createBar(w1_l1, excel, wall[i], san[i]);
            #endregion
            }
            #endregion
            ExportExcel export = new ExportExcel(path, L0_bien, L1_bien, L2_bien, Comment_bien,L0_giua,L1_giua,L2_giua,Comment_giua);
        }
        public double[] z_offset(List<double> L, List<bool> vuotTang, List<double> D, List<double> hVach, double dtc,int n1_)
        {
            double[] z_offset = new double[hVach.Count];
            for (int i = 0; i < hVach.Count; i++)
            {
                if (i == 0) //tầng đầu tiên
                {
                    z_offset[i]=dtc - n1_ * D[i];
                }
                else
                {
                    if (i >= 2)
                        if (vuotTang[i - 2] == true)
                        {
                            z_offset[i]=z_offset[i - 2] + L[i - 2] - hVach[i - 2] - hVach[i - 1] - n1_ * D[i]; continue;
                        }
                    if (vuotTang[i - 1] == true)
                    {
                        z_offset[i]= z_offset[i - 1] + L[i - 1] - hVach[i - 1] - n1_ * D[i+1]; continue;
                    }
                    z_offset[i]=z_offset[i - 1] + L[i - 1] - hVach[i - 1] - n1_ * D[i];
                }
            }
            return z_offset;
        }
        public PointComparePolygonResult CheckPointInWall(XYZ point, Wall wall,double a,double bop)
        {
            Polygon pl = (new WallGeometryInfo(wall)).BottomPolygon;
            Polygon pl1 = CheckGeometry.OffsetPolygon(pl, a, true);
            Polygon pl2 = CheckGeometry.OffsetPolygon(pl1,bop, false);
            return pl2.CheckXYZPointPosition(CheckGeometry.GetProjectPoint(pl2.Plane, point));
        }
        /// <summary>
        /// Xác định cách vẽ ứng với Rebar Point
        /// </summary>
        /// <param name="rbPoint1">rebar Point tầng 1</param>
        /// <param name="rbPoint2">rebar Point tầng 2</param>
        /// <param name="wall2"> wall tầng 2</param>
        /// <param name="a"> khoãng cách tâm thép đến mép</param>
        /// <param name="excel">import excel để lấy n1=40D và đoạn bóp bẻ ke cho phép </param>
        public void DentifyRebar(RebarPoint rbPoint1, RebarPoint rbPoint2, Wall wall2,double a,ImportExcel excel)
        {
            int n1 = 0; int n2 = 0;//list thép ở tầng i và tầng i+1
            double hvach2 = wall2.LookupParameter("Unconnected Height").AsDouble();
            while (n1 < rbPoint1.point.Count || n2 < rbPoint2.point.Count)
            {
                #region "các cây thép wall_i+1 bị thừa, phải cắm mới"
                if (n1 >= rbPoint1.point.Count) 
                {
                    for (int m = n2; m < rbPoint2.point.Count; m++)
                    {
                        XYZ startPoint = GeomUtil.OffsetPoint(rbPoint2.point[m], -XYZ.BasisZ, rbPoint2.z_offset[m]);
                        XYZ endPoint = GeomUtil.OffsetPoint(rbPoint2.point[m], XYZ.BasisZ, excel.n1_ * rbPoint2.D[m]);
                        rbPoint1.addNSPoint(endPoint, startPoint, rbPoint2.type_D[m], wall2);
                    }
                    break;
                }
                #endregion
                #region "các cây thép wall_i bị thừa, phải khóa đầu"
                if (n2 >= rbPoint2.point.Count)
                {
                    for (int m = n1; m < rbPoint1.point.Count; m++)
                        rbPoint1.draw[m] = draw.hook;
                    break;
                }
                #endregion
                #region "xét thanh thép wall1, có nằm trong tiết diện của wall2 ko?"
                if (CheckPointInWall(rbPoint1.point[n1], wall2, a, excel.bop) == PointComparePolygonResult.Outside)
                {
                    rbPoint1.draw[n1] = draw.hook; n1++; continue;
                }
                #endregion
                XYZ vecto = Vecto_XY(rbPoint1.point[n1], rbPoint2.point[n2]);
                double huong = vecto.DotProduct(rbPoint1.normal[n1]);
                #region "phân loại trường hợp vượt tầng"
                if (rbPoint1.L[n1] == 0) //TH tầng 1 ko có thép =sai
                {
                    if (vecto.GetLength() < 10 * mm2f) n2++;
                    n1++;
                    continue;
                }
                else if (rbPoint2.z_offset[n2] >= hvach2) //TH tầng 1 vượt tầng
                {
                    if (rbPoint2.L[n2] == 0)//TH tầng 1 vượt tầng+TH tầng 2 ko có thép = đúng
                    {
                        #region "vẽ theo kiểu vượt tầng, xác định các điểm để vẽ"
                        Polygon bottomWall2 = (new WallGeometryInfo(wall2)).BottomPolygon;
                        XYZ zzStartPoint =  CheckGeometry.GetProjectPoint(bottomWall2.Plane, rbPoint2.point[n2]);
                        double l1 = zzStartPoint.Z - rbPoint1.point[n1].Z;
                        XYZ zzEndPoint = GeomUtil.OffsetPoint(zzStartPoint, XYZ.BasisZ,rbPoint1.L[n1]-l1);
                        rbPoint1.ZigZagPoint(n1, zzStartPoint, zzEndPoint);
                        n1++; n2++; continue;
                        #endregion
                    }
                    else //TH tầng 1 vượt tầng+tầng 2 có các kiểu vẽ khác =sai
                    {
                        n2++; continue;
                    }
                }
                else if (rbPoint2.L[n2] == 0)//TH tầng 1 có thép và ko vượt tầng && tầng 2 lại ko có thép =sai 
                {
                    n2++; continue;
                }
                #endregion
                //Check distance 2 Point and 2 Diameter.
                if (vecto.GetLength() > excel.bop || rbPoint1.D[n1] != rbPoint2.D[n2])
                {
                    //huong >0 thi n2 giu nguyen tiep tuc voi nextRebar_Wall1
                    if (huong > 0) //khóa đầu ở tầng dưới
                    {
                        rbPoint1.draw[n1] = draw.hook;
                        n1++;
                        continue;
                    }
                    else//cắm mới ở tầng trên
                    {
                        #region "cắm mới ở tầng i+1"
                        XYZ startPoint = GeomUtil.OffsetPoint(rbPoint2.point[n2], -XYZ.BasisZ, rbPoint2.z_offset[n2]);
                        XYZ endPoint = GeomUtil.OffsetPoint(rbPoint2.point[n2], XYZ.BasisZ, excel.n1_ * rbPoint2.D[n2]);
                        //đánh dấu lại vào list cắm mới
                        rbPoint1.addNSPoint(endPoint, startPoint, rbPoint2.type_D[n2], wall2);
                        n2++; continue;
                        #endregion
                    }
                }
                else //vẽ zigzag được
                {
                    XYZ startPoint = GeomUtil.OffsetPoint(rbPoint2.point[n2], -XYZ.BasisZ, rbPoint2.z_offset[n2]);
                    XYZ endPoint = GeomUtil.OffsetPoint(rbPoint2.point[n2], XYZ.BasisZ, excel.n1_ * rbPoint2.D[n2]);
                    rbPoint1.ZigZagPoint(n1, startPoint, endPoint);
                    n1++; n2++; continue;
                }
            }
        }
        public void DentifyRebar(RebarPoint rbPoint1, draw draw)
        {
            for (int i = 0; i < rbPoint1.draw.Count; i++)
            {
                rbPoint1.draw[i] = draw;
            }
        }
        public XYZ Vecto_XY(XYZ p1,XYZ p2)
        {
            return new XYZ(p2.X - p1.X, p2.Y - p1.Y, 0);
        }
        /// <summary>
        /// Rebar Point: là 1 list các điểm chứ thông tin vẽ thép
        /// </summary>
        public class RebarPoint
        {
            public int number { get; set; }
            public bool sole_dau { get; set; }
            public bool sole_cuoi { get; set; }
            public List<NSPoint> nsPoint { get; set; }
            public List<double> z_offset { get; set; }
            public List<XYZ> point { get; set; }
            public List<XYZ> zzStartPoint { get; set; }
            public List<XYZ> zzEndPoint { get; set; }
            public List<XYZ> normal { get; set; }
            public List<draw> draw { get; set; }
            public List<double> D { get; set; }
            public List<double> L { get; set; }
            public List<RebarBarType> type_D { get; set; }
            public bool convert(bool sole)
            {
                if (sole==true)
                {
                    return  sole = false;
                }
                else
                {
                    return sole = true;
                }
            }
            public RebarPoint(bool haveStartBar, XYZ startPoint,XYZ normal, int number, double distribution, 
                double D,RebarBarType type_D, double L1, double L2, double z1, double z2,bool sole)
            {
                #region "khai báo"
                this.number = number;
                this.sole_dau = convert(sole);
                this.nsPoint = new List<NSPoint>();
                point = new List<XYZ>();
                this.z_offset = new List<double>();
                zzStartPoint = new List<XYZ>();
                zzEndPoint = new List<XYZ>();
                draw = new List<draw>();
                this.D = new List<double>();
                this.type_D = new List<RebarBarType>();
                this.L = new List<double>();
                this.normal = new List<XYZ>();
                double segment;
                int iMin, iMax;
                #endregion
                if (haveStartBar == true)
                {
                    segment = distribution / ((double)number - 1);
                    iMin = 0;iMax = number - 1;
                }
                else
                {
                    segment = distribution / ((double)number + 1);
                    iMin = 1; iMax = number;
                }
                sole_cuoi = sole;
                for (int i = iMin; i <= iMax; i++)
                {
                    this.sole_cuoi = convert(sole_cuoi);
                    XYZ p = GeomUtil.OffsetPoint(startPoint, normal, segment * i);
                    if (this.sole_cuoi == true)
                    {
                        this.z_offset.Add(z1);
                        point.Add(GeomUtil.OffsetPoint(p, XYZ.BasisZ, z1));
                        this.L.Add(L1);
                    }
                    else
                    {
                        this.z_offset.Add(z2);
                        point.Add(GeomUtil.OffsetPoint(p, XYZ.BasisZ, z2));
                        this.L.Add(L2);
                    }
                    //nếu L=0 ko vẽ
                    draw.Add(StandardBarWall.draw.straight);
                    this.D.Add(D);
                    this.type_D.Add(type_D);
                    this.normal.Add(normal);
                    zzStartPoint.Add(new XYZ(0, 0, 0));
                    zzEndPoint.Add(new XYZ(0, 0, 0));
                }
                //tra lai gia tri sole_cuoi
            }
            public RebarPoint Add(RebarPoint rebarPoint)
            {
                this.number += rebarPoint.number;
                this.z_offset.AddRange(rebarPoint.z_offset);
                point.AddRange(rebarPoint.point);
                zzStartPoint.AddRange(rebarPoint.zzStartPoint);
                zzEndPoint.AddRange(rebarPoint.zzEndPoint);
                normal.AddRange(rebarPoint.normal);
                draw.AddRange(rebarPoint.draw);
                D.AddRange(rebarPoint.D);
                L.AddRange(rebarPoint.L);
                type_D.AddRange(rebarPoint.type_D);
                nsPoint.AddRange(rebarPoint.nsPoint);
                this.sole_cuoi = rebarPoint.sole_cuoi;
                return this;
            }
            public void addNSPoint(XYZ endPoint, XYZ startPoint, RebarBarType rebarType,Wall host)
            {
                nsPoint.Add(new NSPoint(endPoint, startPoint, rebarType, host));
            }
            public class NSPoint
            {
                public XYZ endPoint { get; set; }
                public XYZ startPoint { get; set; }
                public RebarBarType rebarType { get; set; }
                public Wall host { get; set; }
                public NSPoint(XYZ endPoint, XYZ startPoint, RebarBarType rebarType,Wall host)
                {
                    this.endPoint = endPoint;
                    this.startPoint = startPoint;
                    this.rebarType = rebarType;
                    this.host = host;
                }
            }
            public XYZ Vecto_XY(XYZ p1, XYZ p2)
            {
                return new XYZ(p2.X - p1.X, p2.Y - p1.Y, 0);
            }
            public void ZigZagPoint(int index, XYZ zzStartPoint,XYZ zzEndPoint)
            {
                if(Vecto_XY(point[index],zzStartPoint).GetLength()< D[index])
                {
                    draw[index] = StandardBarWall.draw.straight;
                    this.zzEndPoint[index] = zzEndPoint;
                }
                else
                {
                    draw[index] = StandardBarWall.draw.zigzag;
                    this.zzStartPoint[index] = zzStartPoint;
                    this.zzEndPoint[index] = zzEndPoint;
                }
            }
        }
        public List<Rebar> createBar(RebarPoint rbPoint,ImportExcel excel,Wall wall,floorOnWall floor)
        {
            List<Rebar> listBar = new List<Rebar>();
            for (int i = 0; i < rbPoint.point.Count; i++)
            {
                XYZ normal = rbPoint.normal[i];
                XYZ vecZ = XYZ.BasisZ;
                XYZ endPoint1; XYZ endPoint2;
                IList<Curve> curves = new List<Curve>();
                if (rbPoint.L[i] == 0) continue;
                switch (rbPoint.draw[i])
                {
                    case draw.nodraw:
                        continue;
                    case draw.straight:                 
                        curves.Add(Line.CreateBound(rbPoint.point[i], rbPoint.zzEndPoint[i]));
                        break;
                    case draw.zigzag:
                        double lengthBar_wall1 = rbPoint.zzStartPoint[i].Z - rbPoint.point[i].Z;
                        double lengthBar_wall2 = rbPoint.L[i] - lengthBar_wall1;
                        XYZ vectob = Vecto_XY(rbPoint.zzStartPoint[i], rbPoint.point[i]);
                        normal = vecZ.CrossProduct(vectob);
                        double b = vectob.GetLength();
                        double l = lengthBar_wall1 - b * (double)excel.angle_Bar;
                        if (l<=0)
                        {
                            goto case draw.straight;
                        }
                        endPoint1 = GeomUtil.OffsetPoint(rbPoint.point[i], vecZ, l);
                        //endPoint2 = GeomUtil.OffsetPoint(rbPoint.zzStartPoint[i], vecZ, lengthBar_wall2);
                        curves.Add(Line.CreateBound(rbPoint.point[i], endPoint1));
                        curves.Add(Line.CreateBound(endPoint1, rbPoint.zzStartPoint[i]));
                        curves.Add(Line.CreateBound(rbPoint.zzStartPoint[i], rbPoint.zzEndPoint[i]));
                        break;
                    case draw.hook:
                        double hook_Oxy= excel.hook;
                        switch (excel.type_hook)
                        {
                            case "hook1": //hook = 300
                                hook_Oxy = excel.hook;
                                break;
                            case "hook2":
                                hook_Oxy = excel.hook * rbPoint.D[i];
                                break;
                            case "hook3":
                                hook_Oxy = excel.hook - (floor.thickness - excel.hook_offset);
                                break;
                            case "hook4":
                                hook_Oxy = excel.hook * rbPoint.D[i] - (floor.thickness - excel.hook_offset);
                                break;
                        }
                        endPoint1 = new XYZ(rbPoint.point[i].X, rbPoint.point[i].Y, floor.z - excel.hook_offset);
                        endPoint2 = GeomUtil.OffsetPoint(endPoint1, vecZ.CrossProduct(rbPoint.normal[i]), hook_Oxy);
                        curves.Add(Line.CreateBound(rbPoint.point[i], endPoint1));
                        curves.Add(Line.CreateBound(endPoint1, endPoint2));
                        break;
                }
                #region "Vẽ thép và gán partition"
                Rebar rebar = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rbPoint.type_D[i], null, null,
            wall, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                AssignPartition(rebar, wall);
                listBar.Add(rebar);
                #endregion
            }
            #region "vẽ cắm mới"
            for (int i = 0; i < rbPoint.nsPoint.Count; i++)
            {
                double hook_Oz = excel.implant;// đoạn thẳng cắm dọc theo trục XYZ.basisZ tính từ mặt sàn( nằm dưới chân wall) 
                switch (excel.type_implant)
                {
                    case "hook1":
                        hook_Oz = excel.implant;
                        break;
                    case "hook2":
                        hook_Oz = excel.implant * rbPoint.nsPoint[i].rebarType.BarDiameter;
                        break;
                }
                IList<Curve> curves = new List<Curve>();
                XYZ startPoint = GeomUtil.OffsetPoint(rbPoint.nsPoint[i].startPoint, -XYZ.BasisZ, hook_Oz);
                curves.Add(Line.CreateBound(startPoint, rbPoint.nsPoint[i].endPoint));
                #region "Vẽ thép cắm mới và gán partition"
                Rebar rebar = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rbPoint.nsPoint[i].rebarType, null, null,
             rbPoint.nsPoint[i].host, rbPoint.normal[i], curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                AssignPartition(rebar, wall);
                listBar.Add(rebar);
                #endregion
            }
            #endregion
            return listBar;
        }
        public void AssignPartition(Rebar rebar,Wall wall)
        {
            string partition = "";
            partition += wall.LookupParameter("Base Constraint").AsValueString();
            partition += "_" + wall.LookupParameter("Mark").AsString();
            rebar.LookupParameter("Partition").Set(partition);
        }
        public class CombineWall
        {
            #region "check has zigzag between 2 walls
            public draw draw_w1 { get; set; }
            public draw draw_w2 { get; set; }
            public draw draw_l1 { get; set; }
            public draw draw_l2 { get; set; }
            #endregion
            #region "distance zigzag between 2 walls
            public double offset_w1 { get; set; }
            public double offset_w2 { get; set; }
            public double offset_l1 { get; set; }
            public double offset_l2 { get; set; }
            #endregion
            #region "kích thước"
            public Wall wall1 { get; set; }
            public double width1 { get; set; }
            public double length1 { get; set; }
            public double height1 { get; set; }
            public Wall wall2 { get; set; }
            public Line driving2 { get; set; }
            public double width2 { get; set; }
            public double length2 { get; set; }
            public double height2 { get; set; }
            public XYZ sp1 { get; set; }
            public XYZ ep1 { get; set; }
            public double hook1 { get; set; }
            #endregion
            #region "drivingCurve, hướng"
            public Line driving1 { get; set; }
            public XYZ vecX1 { get; set; }
            public XYZ vecY1 { get; set; }
            public XYZ vecX2 { get; set; }
            public XYZ vecY2 { get; set; }

            #endregion
            public CombineWall(Wall wall1, Wall wall2, ImportExcel data)
            {
                this.wall1 = wall1;
                this.wall2 = wall2;
                #region "Mặc định là vẽ thẳng"
                draw_w1 = draw.straight; draw_w2 = draw.straight;
                draw_l1 = draw.straight; draw_l2 = draw.straight;
                double ss = 1 * mm2f;
                #endregion
                #region "driving curve, length, width,height, vecto, sp, ep
                driving1 = DrivingLine(wall1);
                width1 = wall1.WallType.LookupParameter("Width").AsDouble();
                length1 = wall1.LookupParameter("Length").AsDouble();
                height1 = wall1.LookupParameter("Unconnected Height").AsDouble();
                vecX1 = CheckGeometry.GetDirection(driving1);
                vecY1 = XYZ.BasisZ.CrossProduct(vecX1);
                sp1 = GeomUtil.OffsetPoint(driving1.GetEndPoint(0) as XYZ, vecY1, width1 / 2);
                ep1 = GeomUtil.OffsetPoint(driving1.GetEndPoint(1) as XYZ, -vecY1, width1 / 2);


                driving2 = DrivingLine(wall2);
                width2 = wall2.WallType.LookupParameter("Width").AsDouble();
                length2 = wall2.LookupParameter("Length").AsDouble();
                height2 = wall2.LookupParameter("Unconnected Height").AsDouble();
                vecX2 = CheckGeometry.GetDirection(driving2);
                vecY2 = XYZ.BasisZ.CrossProduct(vecX2);
                XYZ sp2 = GeomUtil.OffsetPoint(driving2.GetEndPoint(0), vecY2, width2 / 2);
                XYZ ep2 = GeomUtil.OffsetPoint(driving2.GetEndPoint(1), -vecY2, width2 / 2);
                #endregion
                #region "distance zigzag between 2 walls
                offset_w1 = Distance2Point_vecto_Oxy(sp1, sp2, vecX1);
                offset_l1 = Distance2Point_vecto_Oxy(sp1, sp2, vecY1);
                offset_w2 = Distance2Point_vecto_Oxy(ep1, ep2, vecX1);
                offset_l2 = Distance2Point_vecto_Oxy(ep1, ep2, vecY1);
                #endregion
                #region "check has zigzag between 2 walls
                if (offset_w1 > ss)
                {
                    if (offset_w1 <= data.bop + ss && width1 == width2)
                    {
                        draw_w1 = draw.zigzag;
                    }
                    else
                    {
                        draw_w1 = draw.hook;
                    }
                }
                if (offset_w2 > ss)
                {
                    if (offset_w2 <= data.bop + ss && width1 == width2)
                    {
                        draw_w2 = draw.zigzag;
                    }
                    else
                    {
                        draw_w2 = draw.hook;
                    }
                }
                if (offset_l1 > ss)
                {
                    if (offset_l1 <= data.bop + ss && length1 == length2)
                    {
                        draw_l1 = draw.zigzag;
                    }
                    else
                    {
                        draw_l1 = draw.hook;
                    }
                }
                if (offset_l2 > ss)
                {
                    if (offset_l2 <= data.bop + ss && length1 == length2)
                    {
                        draw_l2 = draw.zigzag;
                    }
                    else
                    {
                        draw_l2 = draw.hook;
                    }
                }
                #endregion
            }
            public CombineWall(Wall wall1)
            {
                this.wall1 = wall1;
                #region "tầng cuối cùng nên vẽ hook"
                draw_w1 = draw.hook; draw_w2 = draw.hook;
                draw_l1 = draw.hook; draw_l2 = draw.hook;
                #endregion
                #region "length, width, Vecto,sp,ep
                driving1 = DrivingLine(wall1);
                width1 = wall1.WallType.LookupParameter("Width").AsDouble();
                length1 = wall1.LookupParameter("Length").AsDouble();
                height1 = wall1.LookupParameter("Unconnected Height").AsDouble();
                vecX1 = CheckGeometry.GetDirection(driving1);
                vecY1 = XYZ.BasisZ.CrossProduct(vecX1);
                sp1 = GeomUtil.OffsetPoint(driving1.GetEndPoint(0), vecY1, width1 / 2);
                ep1 = GeomUtil.OffsetPoint(driving1.GetEndPoint(1), -vecY1, width1 / 2);
                #endregion
            }

            public Line DrivingLine(Element e)
            {
                LocationCurve lc = e.Location as LocationCurve; Curve c = lc.Curve;
                Parameter boP = e.LookupParameter("Base Offset");
                c = GeomUtil.OffsetCurve(c, XYZ.BasisZ, boP.AsDouble());
                List<XYZ> ps = new List<XYZ> { c.GetEndPoint(0), c.GetEndPoint(1) };
                ps.Sort(new ZYXComparer());
                return Line.CreateBound(ps[0], ps[1]);
            }
            public double Distance2Point_vecto_Oxy(XYZ P1, XYZ P2, XYZ vec_fromP1)
            {
                XYZ p11 = GeomUtil.OffsetPoint(P1, vec_fromP1, 1);
                Line line = Line.CreateBound(P1, p11);
                XYZ p12 = CheckGeometry.GetProjectPoint(line, P2);
                return Math.Pow(Math.Pow(p12.X - P1.X, 2) + Math.Pow(p12.Y - P1.Y, 2), 0.5);
            }
        }
        /// <summary> 
        /// xác định các sàn nằm phía trên và chạm vào đỉnh cột
        /// </summary>
        public class floorOnWall
        {
            public double thickness { get; set; }
            public double offset { get; set; }
            public double z { get; set; }
            public floorOnWall(Wall wall,Document doc,Selection sel)
            {
                CombineWall combine = new CombineWall(wall);
                XYZ point0 = GeomUtil.OffsetPoint(combine.sp1,( combine.vecX1-combine.vecY1), 50*mm2f);
                XYZ point1 = GeomUtil.OffsetPoint(combine.ep1,(-combine.vecX1 + combine.vecY1), 50*mm2f);
                //chiếu lên trên
                double heightOffset = combine.height1 / 4;
                point0 = GeomUtil.OffsetPoint(point0, XYZ.BasisZ, heightOffset);
                point1 = GeomUtil.OffsetPoint(point1, XYZ.BasisZ, combine.height1+ heightOffset);
                Outline ol = createOutline(point0, point1);
                BoundingBoxIntersectsFilter filter1 = new BoundingBoxIntersectsFilter(ol);
                ElementIntersectsElementFilter filter2 = new ElementIntersectsElementFilter(wall);
                LogicalAndFilter filter = new LogicalAndFilter(filter1, filter2);
                List<Element> listSan = new FilteredElementCollector(doc).WherePasses(filter).OfClass(typeof(Floor)).ToList();
                Element san;
                if (listSan.Count >0)
                {
                    san = listSan.First();
                    thickness = san.LookupParameter("Thickness").AsDouble();
                    offset = san.LookupParameter("Height Offset From Level").AsDouble();
                    z = san.LookupParameter("Elevation at Top").AsDouble();
                }
                else
                {
                    //ko có sàn
                    thickness = 0;
                    offset =0;
                    z = 0;
                }
            }
        }
        /// <summary>
        /// Create OutLine that don't care Min Point or Max Point
        /// </summary>
        /// <param name="p1">point 1</param>
        /// <param name="p2">point 2</param>
        /// <returns></returns>
        public static Outline createOutline(XYZ p1,XYZ p2)
        {
            XYZ max = new XYZ(Math.Max(p1.X, p2.X), Math.Max(p1.Y, p2.Y), Math.Max(p1.Z, p2.Z));
            XYZ min = new XYZ(Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y), Math.Min(p1.Z, p2.Z));
            return new Outline(min, max);
        }
        /// <summary>
        /// khoãng cách của 2 điểm trong không gian chiếu lên mặt phẳng Oxy
        /// </summary>
        /// <param name="p1">point 1</param>
        /// <param name="p2">point 2</param>
        /// <returns></returns>
        public double distance2Point_Oxy(XYZ p1,XYZ p2)
        {
           return Math.Pow((p1.X - p2.X) * (p1.X - p2.X) + (p1.Y - p2.Y) * (p1.Y - p2.Y), 0.5);
        }
        /// <summary>khử lỗi: do hướng của người vẽ </summary>
        /// <param name="wall"></param>
        /// <returns></returns>
        public Line drivingLine(Element wall)
        {
            LocationCurve lc = wall.Location as LocationCurve; Curve c = lc.Curve;
            Parameter boP = wall.LookupParameter("Base Offset");
            c = GeomUtil.OffsetCurve(c, XYZ.BasisZ, boP.AsDouble());
            List<XYZ> ps = new List<XYZ> { c.GetEndPoint(0), c.GetEndPoint(1) };
            ps.Sort(new ZYXComparer());
            return Line.CreateBound(ps[0], ps[1]);
        }
        /// <summary>lọc ra wall, sắp xếp theo level </summary>
        /// <param name="WallSelected"> list đối tượng được chọn</param>
        /// <returns></returns>
        public List<Wall> sortWallByLevel(Selection WallSelected)
        {
            IList<ElementId> wallListId = WallSelected.GetElementIds() as IList<ElementId>;
            List<Tuple<Wall, double>> eId_z = new List<Tuple<Wall, double>>();
            foreach (ElementId eId in wallListId)
            {
                Element e = doc.GetElement(eId);Wall wall = e as Wall;
                if (wall!=null)
                {
                    double z = (wall.Location as LocationCurve).Curve.GetEndPoint(0).Z;
                    eId_z.Add(new Tuple<Wall, double>(wall, z));
                }
            }
            eId_z = eId_z.OrderBy(j => j.Item2).ToList();
            List<Wall> wallList = new List<Wall>();
            for (int i = 0; i < eId_z.Count; i++)
            {
                wallList.Add(eId_z[i].Item1);
            }
            return wallList;
        }
        public List<Wall> sortWallByLevel(List<Wall> WallSelected)
        {
            List<Tuple<Wall, double>> eId_z = new List<Tuple<Wall, double>>();
            foreach (Wall wall in WallSelected)
            {
                    double z = (wall.Location as LocationCurve).Curve.GetEndPoint(0).Z;
                    eId_z.Add(new Tuple<Wall, double>(wall, z));
            }
            eId_z = eId_z.OrderBy(j => j.Item2).ToList();
            List<Wall> wallList = new List<Wall>();
            for (int i = 0; i < eId_z.Count; i++)
            {
                wallList.Add(eId_z[i].Item1);
            }
            return wallList;
        }
        public class CombineRebar
        {
            public List<double> L0 { get; set; }
            public List<double> L1 { get; set; }
            public List<double> L2 { get; set; }
            public List<string> comment { get; set; }
            public List<bool> vuotTang1 { get; set; }
            public List<bool> vuotTang2 { get; set; }
            public vong Optimize(vong optimize, vong vong)
            {
                double sumL0_vong = 0; double sumL0_optimize = 0;
                if(optimize == null) return vong;
                for (int i = 0; i < vong.L0.Count; i++) 
                {
                    sumL0_optimize += optimize.L0[i];
                    sumL0_vong += vong.L0[i];
                }
                if (sumL0_optimize  > sumL0_vong)
                {
                    return vong;
                }
                else
                {
                    return optimize;
                }
            }
            /// <summary> Phối thép cột vách</summary>
            /// <param name="excel"> import excel </param>
            /// <param name="k">số tầng được phối</param>
            /// <param name="hCot">chiều cao cột/chiều cao tầng</param>
            /// <param name="hSan">chiều dày sàn nằm trên cột</param>
            /// <param name="bt">vùng giới hạn trên</param>
            /// <param name="bd">vùng giới hạn dưới</param>
            /// <param name="D">hiện tại xem D_bien = D_giua</param>
            /// <param name="n1">đoạn nối 40*D nếu coupler thì n1=0 </param>
            /// <param name="n2">khoãng cách giữa 2 đoạn nối </param>
            /// <param name="hook"> đoạn neo vách vào sàn qui định trong data thường là 300</param>
            /// <param name="dtc1"> đoạn thép chờ 1</param>
            /// <param name="dtc2"> đoạn thép chờ 2</param>
            public CombineRebar(ImportExcel excel, List<double> hCot, List<floorOnWall> san, List<double> bt, List<double> bd,
                List<double> D, double dtc1, double dtc2)
            {
                #region "khai báo"
                double Lsum = 11700 * mm2f;//input
                int k = excel.k; int n1 = excel.n1_; int n2 = excel.n2_; double hook = excel.hook;//input
                double Lmin = excel.Lmin; double Lmax = excel.Lmax; double step = -50 * mm2f;//input
                List<double> Ht = new List<double>(); //calculate
                List<double> kc = new List<double>();//caculate
                L1 = new List<double>(); //calculate - OutPut
                L2 = new List<double>(); //calculate - OutPut
                L0 = new List<double>(); //calculate - OutPut
                comment = new List<string>();//calculate - OutPut
                List<double> d1 = new List<double>(); //calculate
                List<double> d2 = new List<double>(); //calculate
                List<double> d11 = new List<double>(); //calculate
                List<double> d22 = new List<double>(); //calculate 
                vuotTang1 = new List<bool>();//calculate
                vuotTang2 = new List<bool>();//calculate          
                List<bool> hoiQuy = new List<bool>(0); bool hq = false;//caluculate
                List<vong> vong = new List<vong>();
                int i, j1, j2;//tương ứng với wall[i] L1[j1] L2[j2]
                for (i = 0; i < hCot.Count; i++)
                {
                    kc.Add((n1 + n2) * D[i]);
                    L1.Add(0); L2.Add(0); L0.Add(0); comment.Add("");
                    d1.Add(0); d11.Add(0); d2.Add(0); d22.Add(0);
                    vuotTang1.Add(false); vuotTang2.Add(false);
                    hoiQuy.Add(false);
                    vong.Add(null);
                    if (i < hCot.Count - 1) Ht.Add(hCot[i] + bt[i + 1] - bt[i]);
                    else Ht.Add(hCot[i] - bt[i]);
                }
                d1[0] = dtc1 - bt[0]; d2[0] = dtc2 - bt[0];//calculate d1[0],d2[0]
                List<double> module1 = new List<double> { 3900 * mm2f, 5850 * mm2f, 2925 * mm2f }; //calculate
                List<double> module2 = new List<double>(module1); //calculate
                for (double li = Lmax; li >= Lmin; li += step) module2.Add(li);
                #endregion
                for (i = 0; i < Ht.Count; i++)
                {
                    #region "Hồi quy-i,L0,comment"
                    if (i >= 1)
                    {
                        hq = hoiQuy[i - 1];
                        if (hq == true)
                            if (i == 1)
                            {
                                TaskDialog.Show("Thông báo", "không có trường hợp phối thép nào thỏa mãn!");
                                break;
                            }
                            else
                            {
                                i = i - 2;
                                L0 = new List<double>(vong[i].L0);
                                comment = new List<string>(vong[i].comment);
                                vuotTang1 = new List<bool>(vong[i].vuotTang1);
                                vuotTang2 = new List<bool>(vong[i].vuotTang2);
                            }
                    }
                    #endregion
                    #region "wall[i]=wall[end]"
                    if (i == Ht.Count - 1)
                    {
                        //google tùy kiểu hook
                        L1[i] = (hook + n1 * D[i]) + (hCot[i] + san[i].offset - san[i].thickness) - bt[i] - d1[i];
                        L2[i] = (hook + n1 * D[i]) + (hCot[i] + san[i].offset - san[i].thickness) - bt[i] - d2[i];
                        d1[i] = 0; d2[i] = 0; d11[i] = 0; d22[i] = 0;
                        L0[i] = Lsum - L1[i] - L2[i];
                        break;
                    }
                    #endregion
                    #region "create Module0 từ L0(đoạn thép dư) và loại L0[t]=0"
                    List<double> module0 = new List<double>();
                    if (i > 0 && D[i] == D[i - 1]) for (int t = Math.Max(1, i - k); t < i; t++) module0.Add(L0[t]);
                    #endregion
                    #region "Hồi quy=true - nhảy đến cách chọn L1,L2"
                    if (hq == true) //nếu hồi quy
                    {
                        if (vong[i].cach == cach.c11 || vong[i].cach == cach.c12) goto cach1;
                        if (vong[i].cach == cach.c21 || vong[i].cach == cach.c22) goto cach2;
                        if (vong[i].cach == cach.c31 || vong[i].cach == cach.c32) goto cach3;
                    }
                #endregion
                cach1:
                    #region "Cách 1: chọn L1,L2 trong Module0=List L0+Module1"
                    List<double> module_c1 = new List<double>(module0); module_c1.AddRange(module1);
                    for (j1 = 0; j1 < module_c1.Count; j1++)
                    {
                        if (module_c1[j1] == 0) continue;
                        for (j2 = j1; j2 < module_c1.Count; j2++)
                        {
                            if (module_c1[j2] == 0) continue;
                            #region "Hồi quy-start"
                            if (hq == true)
                            {
                                j1 = vong[i].j1;
                                j2 = vong[i].j2;
                                if (vong[i].cach == cach.c11) goto endC11;
                                if (vong[i].cach == cach.c12) goto endC12;
                            }
                            #endregion
                            if (j1 < module0.Count && j1 == j2) continue;
                            L1[i] = module_c1[j1]; L2[i] = module_c1[j2];
                            #region "Check-c11"
                            d1[i + 1] = d1[i] + (L1[i] - Ht[i]) - n1 * D[i];
                            d2[i + 1] = d2[i] + (L2[i] - Ht[i]) - n1 * D[i];
                            //xét có vượt tầng hay ko?
                            if (i - 1 >= 0)
                            {
                                if (vuotTang1[i - 1] == true)
                                {
                                    L1[i] = 0;
                                    d1[i + 1] = d11[i];
                                }
                                if (vuotTang2[i - 1] == true)
                                {
                                    L2[i] = 0;
                                    d2[i + 1] = d22[i];
                                }
                            }
                            d11[i + 1] = d1[i + 1] - Ht[i + 1];
                            d22[i + 1] = d2[i + 1] - Ht[i + 1];
                            if (Math.Abs(d1[i + 1] - d2[i + 1]) < kc[i + 1] || d1[i + 1] < n1 * D[i + 1] || d2[i + 1] < n1 * D[i + 1]) goto endC11;
                            if (L1[i] == 0) vuotTang1[i] = false;
                            if (L2[i] == 0) vuotTang2[i] = false;
                            if (excel.vuotTang == "yes" && i + 2 <= Ht.Count - 1)
                            {
                                if (d11[i + 1] > -bd[i + 1] - bt[i + 2])
                                {
                                    if (d11[i + 1] < n1*D[i + 2]) goto endC11;
                                    vuotTang1[i] = true;
                                }
                                if (d22[i + 1] > -bd[i + 1] - bt[i + 2])
                                {
                                    if (d22[i + 1] < n1*D[i + 2]) goto endC11;
                                    vuotTang2[i] = true;
                                }
                            }
                            else
                            {
                                if (d11[i + 1] > -bd[i + 1]) goto endC11;
                                if (d22[i + 1] > -bd[i + 1]) goto endC11;
                            }
                            #endregion
                            hoiQuy[i] = false;
                            vong[i] = new vong(i, j1, j2, cach.c11, L0, comment, vuotTang1, vuotTang2);
                            #region"Result-C11"
                            if (j1 < module0.Count)
                            {
                                if (comment[i - (module0.Count - j1)] != "") comment[i - (module0.Count - j1)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j1)] += "tận dụng L" + Math.Round(L0[i - (module0.Count - j1)] * f2mm / 5) * 5 + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j1)] * f2mm / 5) * 5 + " từ tầng " + (i - (module0.Count - j1));
                                L0[i - (module0.Count - j1)] = 0;
                            }
                            if (j2 < module0.Count)
                            {
                                if (comment[i - (module0.Count - j2)] != "") comment[i - (module0.Count - j2)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j2)] += "tận dụng L" + L0[i - (module0.Count - j2)] + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j2)] / 5) * 5 + " từ tầng " + (i - (module0.Count - j2));
                                L0[i - (module0.Count - j2)] = 0;
                            }
                            #endregion
                            L0[i] = 0;
                            goto Next_i;
                        endC11:
                            L1[i] = module_c1[j2]; L2[i] = module_c1[j1];
                            #region "Check-c12"
                            d1[i + 1] = d1[i] + (L1[i] - Ht[i]) - n1*D[i];
                            d2[i + 1] = d2[i] + (L2[i] - Ht[i]) - n1*D[i];
                            if (i - 1 >= 0)
                            {
                                if (vuotTang1[i - 1] == true)
                                {
                                    L1[i] = 0;
                                    d1[i + 1] = d11[i];
                                }
                                if (vuotTang2[i - 1] == true)
                                {
                                    L2[i] = 0;
                                    d2[i + 1] = d22[i];
                                }
                            }
                            d11[i + 1] = d1[i + 1] - Ht[i + 1];
                            d22[i + 1] = d2[i + 1] - Ht[i + 1];
                            if (Math.Abs(d1[i + 1] - d2[i + 1]) < kc[i + 1] || d1[i + 1] < n1 * D[i + 1] || d2[i + 1] < n1 * D[i + 1]) goto endC12;
                            if (L1[i] == 0) vuotTang1[i] = false;
                            if (L2[i] == 0) vuotTang2[i] = false;
                            if (excel.vuotTang == "yes" && i + 2 <= Ht.Count - 1)
                            {
                                if (d11[i + 1] > -bd[i + 1] - bt[i + 2])
                                {
                                    if (d11[i + 1] < n1*D[i + 2]) goto endC12;
                                    vuotTang1[i] = true;
                                }
                                if (d22[i + 1] > -bd[i + 1] - bt[i + 2])
                                {
                                    if (d22[i + 1] < n1*D[i + 2]) goto endC12;
                                    vuotTang2[i] = true;
                                }
                            }
                            else
                            {
                                if (d11[i + 1] > -bd[i + 1]) goto endC12;
                                if (d22[i + 1] > -bd[i + 1]) goto endC12;
                            }
                            #endregion
                            hoiQuy[i] = false;
                            vong[i] = new vong(i, j1, j2, cach.c12, L0, comment, vuotTang1, vuotTang2);
                            #region "Result - C12"
                            if (module0.Count - j1 > 0)
                            {
                                if (comment[i - (module0.Count - j1)] != "") comment[i - (module0.Count - j1)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j1)] += "tận dụng L" + Math.Round(L0[i - (module0.Count - j1)] * f2mm / 5) * 5 + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j1)] * f2mm / 5) * 5 + " từ tầng " + (i - (module0.Count - j1));
                                L0[i - (module0.Count - j1)] = 0;
                            }
                            if (module0.Count - j2 > 0)
                            {
                                if (comment[i - (module0.Count - j2)] != "") comment[i - (module0.Count - j2)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j2)] += "tận dụng L" + Math.Round(L0[i - (module0.Count - j2)] / 5) / 5 + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j2)] / 5) * 5 + " từ tầng " + (i - (module0.Count - j2));
                                L0[i - (module0.Count - j2)] = 0;
                            }
                            #endregion
                            L0[i] = 0;
                            goto Next_i;
                        endC12:;
                        }
                    }
                #endregion
                cach2:
                    #region "Cách 2: chọn L1,L2 trong Module0=List L0+Module2 mà L0 trong Modul1"
                    List<double> module_c2 = new List<double>(module0); module_c2.AddRange(module2);
                    j1 = 0;
                    for (j1 = 0; j1 < module1.Count; j1++)
                    {
                        if (module_c2[j1] == 0) continue;
                        for (j2 = 0; j2 < module_c2.Count; j2++)
                        {
                            if (module_c2[j2] == 0) continue;
                            #region "hoi quy-start"
                            if (hq == true)
                            {
                                j1 = vong[i].j1;
                                j2 = vong[i].j2;
                                if (vong[i].cach == cach.c21) goto endC21;
                                if (vong[i].cach == cach.c22) goto endC22;
                            }
                            #endregion
                            L1[i] = module_c2[j2]; L2[i] = Lsum - module1[j1] - L1[i]; if (L2[i] < Lmin) goto endC21;
                            #region "Check-c21"
                            d1[i + 1] = d1[i] + (L1[i] - Ht[i]) - n1*D[i];
                            d2[i + 1] = d2[i] + (L2[i] - Ht[i]) - n1 * D[i];
                            d11[i + 1] = d1[i + 1] - Ht[i + 1];
                            d22[i + 1] = d2[i + 1] - Ht[i + 1];
                            if (Math.Abs(d1[i + 1] - d2[i + 1]) < kc[i + 1] || d1[i + 1] < n1 * D[i + 1] || d2[i + 1] < n1 * D[i + 1]) goto endC21;
                            if (i + 2 <= Ht.Count - 1)
                            {
                                if (d11[i + 1] > -bd[i + 1] - bt[i + 2]) goto endC21;
                                if (d22[i + 1] > -bd[i + 1] - bt[i + 2]) goto endC21;
                            }
                            else
                            {
                                if (d11[i + 1] > -bd[i + 1]) goto endC21;
                                if (d22[i + 1] > -bd[i + 1]) goto endC21;
                            }
                            #endregion
                            hoiQuy[i] = false;
                            vong[i] = new vong(i, j1, j2, cach.c21, L0, comment, vuotTang1, vuotTang2);
                            #region "Result-c21"
                            if (module0.Count - j2 > 0)
                            {
                                if (comment[i - (module0.Count - j2)] != "") comment[i - (module0.Count - j2)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j2)] += "tận dụng L" + Math.Round(L0[i - (module0.Count - j2)] * f2mm / 5) * 5 + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j2)] * f2mm / 5) * 5 + " từ tầng " + (i - (module0.Count - j2));
                                L0[i - (module0.Count - j2)] = 0;
                            }
                            #endregion
                            L0[i] = module1[j1];
                            goto Next_i;
                        endC21:
                            L2[i] = module_c2[j2]; L1[i] = Lsum - module1[j1] - L2[i]; if (L1[i] < Lmin) goto endC22;
                            #region "Check-c22"
                            d1[i + 1] = d1[i] + (L1[i] - Ht[i]) - n1 * D[i];
                            d2[i + 1] = d2[i] + (L2[i] - Ht[i]) - n1 * D[i];
                            d11[i + 1] = d1[i + 1] - Ht[i + 1];
                            d22[i + 1] = d2[i + 1] - Ht[i + 1];
                            if (Math.Abs(d1[i + 1] - d2[i + 1]) < kc[i + 1] || d1[i + 1] < n1 * D[i + 1] || d2[i + 1] < n1 * D[i + 1]) goto endC22;
                            if (i + 2 <= Ht.Count - 1)
                            {
                                if (d11[i + 1] > -bd[i + 1] - bt[i + 2]) goto endC22;
                                if (d22[i + 1] > -bd[i + 1] - bt[i + 2]) goto endC22;
                            }
                            else
                            {
                                if (d11[i + 1] > -bd[i + 1]) goto endC22;
                                if (d22[i + 1] > -bd[i + 1]) goto endC22;
                            }
                            #endregion
                            hoiQuy[i] = false;
                            vong[i] = new vong(i, j1, j2, cach.c22, L0, comment, vuotTang1, vuotTang2);
                            #region "Result-c22"
                            if (module0.Count - j2 > 0)
                            {
                                if (comment[i - (module0.Count - j2)] != "") comment[i - (module0.Count - j2)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j2)] += "tận dụng L" + Math.Round(L0[i - (module0.Count - j2)] * f2mm / 5) * 5 + " cho tầng" + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j2)] * f2mm / 5) * 5 + " từ tầng" + (i - (module0.Count - j2));
                                L0[i - (module0.Count - j2)] = 0;
                            }
                            #endregion
                            L0[i] = module1[j1];
                            goto Next_i;
                        endC22:;
                        }
                    }
                #endregion
                cach3:
                    #region "Cách 3: chọn L1,L2,L0 trong Module0=List L0+Module2 "
                    List<double> module_c3 = new List<double>(module0); module_c3.AddRange(module2);
                    for (j1 = 0; j1 < module_c3.Count; j1++)
                    {
                        if (module_c3[j1] == 0) continue;
                        for (j2 = j1; j2 < module_c3.Count; j2++)
                        {
                            if (module_c3[j2] == 0) continue;
                            #region "hoi quy-start"
                            if (hq == true)
                            {
                                j1 = vong[i].j1;
                                j2 = vong[i].j2;
                                if (vong[i].cach == cach.c31) goto endC31;
                                if (vong[i].cach == cach.c32) goto endC32;
                            }
                            #endregion
                            if (j1 < module0.Count && j1 == j2) continue;
                            L1[i] = module_c3[j1]; L2[i] = module_c3[j2];
                            #region "Check-c31"
                            d1[i + 1] = d1[i] + (L1[i] - Ht[i]) - n1 * D[i];
                            d2[i + 1] = d2[i] + (L2[i] - Ht[i]) - n1 * D[i];
                            if (i - 1 >= 0)
                            {
                                if (vuotTang1[i - 1] == true)
                                {
                                    L1[i] = 0;
                                    d1[i + 1] = d11[i];
                                }
                                if (vuotTang2[i - 1] == true)
                                {
                                    L2[i] = 0;
                                    d2[i + 1] = d22[i];
                                }
                            }
                            d11[i + 1] = d1[i + 1] - Ht[i + 1];
                            d22[i + 1] = d2[i + 1] - Ht[i + 1];
                            if (Math.Abs(d1[i + 1] - d2[i + 1]) < kc[i + 1] || d1[i + 1] < n1 * D[i + 1] || d2[i + 1] < n1 * D[i + 1]) goto endC31;
                            if (L1[i] == 0) vuotTang1[i] = false;
                            if (L2[i] == 0) vuotTang2[i] = false;
                            if (excel.vuotTang == "yes" && i + 2 <= Ht.Count - 1)
                            {
                                if (d11[i + 1] > -bd[i + 1] - bt[i + 2])
                                {
                                    if (d11[i + 1] < n1 * D[i + 2]) goto endC31;
                                    vuotTang1[i] = true;
                                }
                                if (d22[i + 1] > -bd[i + 1] - bt[i + 2])
                                {
                                    if (d22[i + 1] < n1 * D[i + 2]) goto endC31;
                                    vuotTang2[i] = true;
                                }
                            }
                            else
                            {
                                if (d11[i + 1] > -bd[i + 1]) goto endC31;
                                if (d22[i + 1] > -bd[i + 1]) goto endC31;
                            }
                            #endregion
                            hoiQuy[i] = false;
                            vong[i] = new vong(i, j1, j2, cach.c31, L0, comment, vuotTang1, vuotTang2);
                            #region "Result-c31"
                            if (module0.Count - j1 > 0)
                            {
                                if (comment[i - (module0.Count - j1)] != "") comment[i - (module0.Count - j1)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j1)] += "tận dụng L" + Math.Round(L0[i - (module0.Count - j1)] * f2mm / 5) * 5 + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j1)] * f2mm / 5, 0) * 5 + " từ tầng " + (i - (module0.Count - j1));
                                L0[i - (module0.Count - j1)] = 0;
                            }
                            if (module0.Count - j2 > 0)
                            {
                                if (comment[i - (module0.Count - j2)] != "") comment[i - (module0.Count - j2)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j2)] += "tận dụng L" + L0[i - (module0.Count - j2)] + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j2)] / 5) * 5 + " từ tầng " + (i - (module0.Count - j2));
                                L0[i - (module0.Count - j2)] = 0;
                            }
                            #endregion
                            L0[i] = Lsum - L1[i] - L2[i];
                            goto Next_i;
                        endC31:
                            L1[i] = module_c3[j2]; L2[i] = module_c3[j1];
                            #region "Check-c32"
                            d1[i + 1] = d1[i] + (L1[i] - Ht[i]) - n1 * D[i];
                            d2[i + 1] = d2[i] + (L2[i] - Ht[i]) - n1 * D[i];
                            if (i - 1 >= 0)
                            {
                                if (vuotTang1[i - 1] == true)
                                {
                                    L1[i] = 0;
                                    d1[i + 1] = d11[i];
                                }
                                if (vuotTang2[i - 1] == true)
                                {
                                    L2[i] = 0;
                                    d2[i + 1] = d22[i];
                                }
                            }
                            d11[i + 1] = d1[i + 1] - Ht[i + 1];
                            d22[i + 1] = d2[i + 1] - Ht[i + 1];
                            if (Math.Abs(d1[i + 1] - d2[i + 1]) < kc[i + 1] || d1[i + 1] < n1 * D[i + 1] || d2[i + 1] < n1 * D[i + 1]) goto endC32;
                            if (L1[i] == 0) vuotTang1[i] = false;
                            if (L2[i] == 0) vuotTang2[i] = false;
                            if (excel.vuotTang == "yes" && i + 2 <= Ht.Count - 1)
                            {
                                if (d11[i + 1] > -bd[i + 1] - bt[i + 2])
                                {
                                    if (d11[i + 1] < n1 * D[i + 2]) goto endC32;
                                    vuotTang1[i] = true;
                                }
                                if (d22[i + 1] > -bd[i + 1] - bt[i + 2])
                                {
                                    if (d22[i + 1] < n1 * D[i + 2]) goto endC32;
                                    vuotTang2[i] = true;
                                }
                            }
                            else
                            {
                                if (d11[i + 1] > -bd[i + 1]) goto endC32;
                                if (d22[i + 1] > -bd[i + 1]) goto endC32;
                            }
                            #endregion
                            hoiQuy[i] = false;
                            vong[i] = new vong(i, j1, j2, cach.c32, L0, comment, vuotTang1, vuotTang2);
                            #region "Result-c32"
                            if (module0.Count - j1 > 0)
                            {
                                if (comment[i - (module0.Count - j1)] != "") comment[i - (module0.Count - j1)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j1)] += "tận dụng L" + Math.Round(L0[i - (module0.Count - j1)] * f2mm / 5) * 5 + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j1)] * f2mm / 5, 0) * 5 + " từ tầng " + (i - (module0.Count - j1));
                                L0[i - (module0.Count - j1)] = 0;
                            }
                            if (module0.Count - j2 > 0)
                            {
                                if (comment[i - (module0.Count - j2)] != "") comment[i - (module0.Count - j2)] += "\n";
                                if (comment[i] != "") comment[i] += "\n";
                                comment[i - (module0.Count - j2)] += "tận dụng L" + L0[i - (module0.Count - j2)] + " cho tầng " + i;
                                comment[i] = "tái sử dụng L" + Math.Round(L0[i - (module0.Count - j2)] / 5) * 5 + " từ tầng " + (i - (module0.Count - j2));
                                L0[i - (module0.Count - j2)] = 0;
                            }
                            #endregion
                            L0[i] = Lsum - L1[i] - L2[i];
                            goto Next_i;
                        endC32:;
                        }
                    }
                    #endregion
                    //tầng[i] này ko ra duoc ket qua cần hồi quy lại phương án ở tầng [i+1]
                    hoiQuy[i] = true;
                //TaskDialog.Show("HQ", "Hồi quy 1 lần");
                Next_i:;
                }
            }
            public enum cach
            {
                c11,
                c12,
                c21,
                c22,
                c31,
                c32
            }
            public class vong
            {
                public int i { get; set; }
                public int j1 { get; set; }
                public int j2 { get; set; }
                public cach cach { get; set; }
                public List<double> L0 { get; set; }
                public List<string> comment { get; set; }
                public List<bool> vuotTang1 { get; set; }
                public List<bool> vuotTang2 { get; set; }
                public vong(int i, int j1, int j2, cach cach, List<double> L0, List<string> comment, List<bool>vuotTang1, List<bool> vuotTang2)
                {
                    this.i = i;
                    this.j1 = j1;
                    this.j2 = j2;
                    this.L0 = new List<double>(L0);
                    this.cach = cach;
                    this.comment = new List<string>(comment);
                    this.vuotTang1 = new List<bool>(vuotTang1);
                    this.vuotTang2 = new List<bool>(vuotTang2);
                }
            }
        }
    }
}