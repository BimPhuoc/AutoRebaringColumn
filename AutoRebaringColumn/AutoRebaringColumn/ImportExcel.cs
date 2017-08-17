using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Structure;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices.ComTypes;
using AutoRebaringColumn;
using System.IO;
using System.Xml;

namespace DataExcel
{
    public class ImportExcel
    {
        #region "convert unit"
        const double f2m = 0.3048;
        const double f2mm = f2m * 1000;
        const double mm2f = 1 / f2mm;
        #endregion
        #region "thông tin chung"
        public int colCount { get; set; }
        public string mark { get; set; }
        public double dtc1_bien { get; set; }
        public double dtc2_bien { get; set; }
        public double dtc1_giua { get; set; }
        public double dtc2_giua { get; set; }
        public double hook { get; set; }
        public double hook_offset { get; set; }
        public double bop { get; set; }
        public double ss { get; set; }
        public double Lmin { get; set; }
        public double Lmax { get; set; }
        public int angle_Bar { get; set; }
        public int tl { get; set; }
        public double abv { get; set; }
        public int n1_ { get; set; }
        public int n2_ { get; set; }
        public int k { get; set; }
        public string vuotTang { get; set; }
        public string start_level { get; set; }
        public string end_level { get; set; }
        public string tan_dung { get; set; }
        public double chenh_lechPT { get; set; }
        public int chenh_lechS { get; set; }
        public string type_hook { get; set; }
        public string type_implant { get; set; }
        public double implant { get; set; }
        #endregion
        #region "Set Parameter"
        public string name_para1 { get; set; }
        public string name_para2 { get; set; }
        public string name_para3 { get; set; }
        public string name_para4 { get; set; }
        public List<object> para1 { get; set; }
        public List<object> para2 { get; set; }
        public List<object> para3 { get; set; }
        public List<object> para4 { get; set; }
        #endregion
        #region "Standard Rebar"
        public int LevelCount { get; set; }
        public List<string> Level { get; set; }
        public List<double> bt { get; set; }
        public List<double> bd { get; set; }
        public List<string> D { get; set; }
        public List<int> number_A { get; set; }
        public List<int> number_B { get; set; }
        #endregion
        #region "Stirrup Rebar"
        public List<string> SCDaiNgoai { get; set; }
        public List<string> DKDaiNgoai { get; set; }
        public List<string> SCDaiBien { get; set; }
        public List<string> DKDaiBien { get; set; }
        public List<string> SCDaiCon { get; set; }
        public List<string> DKDaiCon { get; set; }
        public List<double> spacing1_con { get; set; }
        public List<double> spacing2_con { get; set; }
        public List<double> spacing3_bao { get; set; }
        public List<double> spacing4_bao { get; set; }
        public List<double> spacing5_bien { get; set; }
        public List<double> spacing6_bien { get; set; }
        public List<string> daiNgang { get; set; }

        #endregion
        public ImportExcel(int colCount, string excelPath,Document doc)
        {
            this.colCount = colCount;
            ExcelFile ex = new ExcelFile(excelPath);
            Workbook wb = ex.Workbook;
            wb.Save();
            Worksheet sheet = ex.Workbook.Worksheets["WallRebar"];
            #region "thông tin chung"
            mark = (string)sheet.Range["mark"].Value;
            dtc1_bien = (double)sheet.Range["dtc1_bien"].Value * mm2f;
            dtc2_bien = (double)sheet.Range["dtc2_bien"].Value * mm2f;
            dtc1_giua = (double)sheet.Range["dtc1_giua"].Value * mm2f;
            dtc2_giua = (double)sheet.Range["dtc2_giua"].Value * mm2f;
            hook = (double)sheet.Range["hook"].Value * mm2f;
            hook_offset = (double)sheet.Range["hook_offset"].Value * mm2f;
            bop = (double)sheet.Range["bop"].Value * mm2f;
            ss = (double)sheet.Range["ss"].Value * mm2f;
            Lmin = (double)sheet.Range["Lmin"].Value * mm2f;
            Lmax = (double)sheet.Range["Lmax"].Value * mm2f;
            angle_Bar = (int)sheet.Range["angle_Bar"].Value;
            tl = (int)sheet.Range["tl"].Value;
            abv = (double)sheet.Range["abv"].Value * mm2f;
            n1_ = (int)sheet.Range["n1_"].Value;
            n2_ = (int)sheet.Range["n2_"].Value;
            k = (int)sheet.Range["k"].Value;
            vuotTang = (string)sheet.Range["vuotTang"].Value;
            start_level= (string)sheet.Range["start_level"].Value;
            end_level = (string)sheet.Range["end_level"].Value;
            tan_dung = (string)sheet.Range["tan_dung"].Value;
            chenh_lechPT = (double)sheet.Range["chenh_lechPT"].Value;
            chenh_lechS = (int)sheet.Range["chenh_lechS"].Value;
            type_hook = (string)sheet.Range["type_hook"].Value;
            type_implant = (string)sheet.Range["type_implant"].Value;
            implant = (double)sheet.Range["implant"].Value*mm2f;
            #endregion
            #region "name and list Parameter"
            name_para1 = sheet.Range["name_para1"].Value;
            name_para2 = sheet.Range["name_para2"].Value;
            name_para3 = sheet.Range["name_para3"].Value;
            name_para4 = sheet.Range["name_para4"].Value;
            para1 = new List<object>();
            para2 = new List<object>();
            para3 = new List<object>();
            para4 = new List<object>();
            #endregion
            #region "khai báo List"
            Level = new List<string>();
            bt = new List<double>();
            bd = new List<double>();
            D = new List<string>();
            number_A = new List<int>();
            number_B = new List<int>();
            SCDaiNgoai = new List<string>();
            DKDaiNgoai = new List<string>();
            SCDaiBien = new List<string>();
            DKDaiBien = new List<string>();
            SCDaiCon = new List<string>();
            DKDaiCon = new List<string>();
            spacing1_con = new List<double>();
            spacing2_con = new List<double>();
            spacing3_bao = new List<double>();
            spacing4_bao = new List<double>();
            spacing5_bien = new List<double>();
            spacing6_bien = new List<double>();
            daiNgang = new List<string>();
            #endregion
            for (int i = 0; i < colCount; i++)
            {
                #region "standard rebar"
                Level.Add((string)sheet.Range["Level"].Offset[i].Value);
                bt.Add((double)sheet.Range["bt"].Offset[i].Value * mm2f);
                bd.Add((double)sheet.Range["bd"].Offset[i].Value * mm2f);
                D.Add((string)sheet.Range["D"].Offset[i].Value);
                number_A.Add((int)sheet.Range["number_A"].Offset[i].Value);
                number_B.Add((int)sheet.Range["number_B"].Offset[i].Value);
                #endregion
                #region "stirrup rebar"
                SCDaiNgoai.Add((string)sheet.Range["SCDaiNgoai"].Offset[i].Value);
                DKDaiNgoai.Add((string)sheet.Range["DKDaiNgoai"].Offset[i].Value);
                SCDaiBien.Add((string)sheet.Range["SCDaiBien"].Offset[i].Value);
                DKDaiBien.Add((string)sheet.Range["DKDaiBien"].Offset[i].Value);
                SCDaiCon.Add((string)sheet.Range["SCDaiCon"].Offset[i].Value);
                DKDaiCon.Add((string)sheet.Range["DKDaiCon"].Offset[i].Value);
                spacing1_con.Add((double)sheet.Range["spacing1_con"].Offset[i].Value * mm2f);
                spacing2_con.Add((double)sheet.Range["spacing2_con"].Offset[i].Value * mm2f);
                spacing3_bao.Add((double)sheet.Range["spacing3_bao"].Offset[i].Value * mm2f);
                spacing4_bao.Add((double)sheet.Range["spacing4_bao"].Offset[i].Value * mm2f);
                spacing5_bien.Add((double)sheet.Range["spacing5_bien"].Offset[i].Value * mm2f);
                spacing6_bien.Add((double)sheet.Range["spacing6_bien"].Offset[i].Value * mm2f);
                daiNgang.Add((string)sheet.Range["daiNgang"].Offset[i].Value);

                #endregion
                #region "Set Parameter"
                para1.Add(sheet.Range["para1"].Offset[i].Value);
                para2.Add(sheet.Range["para2"].Offset[i].Value);
                para3.Add(sheet.Range["para3"].Offset[i].Value);
                para4.Add(sheet.Range["para4"].Offset[i].Value);
                #endregion
            }
            LevelCount = Level.Count;
        }

        public ImportExcel(string wallMark, string xmlPath)
        {
            this.mark = "";
            if (!File.Exists(xmlPath))
            {
                TaskDialog.Show("Revit","Can not find xmlPath!"); return;
            }
            XmlReader reader = XmlReader.Create(xmlPath);
            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        switch (reader.Name)
                        {
                            case "WallRebarInFoCollection":
                                {
                                    this.hook = double.Parse(reader.GetAttribute("hook"));
                                    this.hook_offset = double.Parse(reader.GetAttribute("hook_offset"));
                                    this.bop = double.Parse(reader.GetAttribute("bop"));
                                    this.ss = double.Parse(reader.GetAttribute("ss"));
                                    this.Lmin = double.Parse(reader.GetAttribute("Lmin"));
                                    this.Lmax = double.Parse(reader.GetAttribute("Lmax"));
                                    this.angle_Bar = int.Parse(reader.GetAttribute("angle_Bar"));
                                    this.tl = int.Parse(reader.GetAttribute("tl"));
                                    this.abv = double.Parse(reader.GetAttribute("abv"));
                                    this.n1_ = int.Parse(reader.GetAttribute("n1_"));
                                    this.n2_ = int.Parse(reader.GetAttribute("n2_"));
                                    this.k = int.Parse(reader.GetAttribute("k"));
                                    this.vuotTang = reader.GetAttribute("vuotTang");
                                    this.tan_dung = reader.GetAttribute("tan_dung");
                                    this.chenh_lechPT = int.Parse(reader.GetAttribute("chenh_lechPT"));
                                    this.chenh_lechS = int.Parse(reader.GetAttribute("chenh_lechS"));
                                    this.type_hook = reader.GetAttribute("type_hook");
                                    this.type_implant = reader.GetAttribute("type_implant");
                                    this.implant = double.Parse(reader.GetAttribute("implant"));
                                    break;
                                }
                            case "WallRebarInfo":
                                {
                                    if (this.mark == wallMark) goto L1;
                                    string s= reader.GetAttribute("mark");
                                    if (s != wallMark) break;
                                    
                                    this.dtc1_bien = double.Parse(reader.GetAttribute("dtc1_bien"));
                                    this.dtc2_bien = double.Parse(reader.GetAttribute("dtc2_bien"));
                                    this.dtc1_giua = double.Parse(reader.GetAttribute("dtc1_giua"));
                                    this.dtc2_giua = double.Parse(reader.GetAttribute("dtc2_giua"));
                                    this.LevelCount = int.Parse(reader.GetAttribute("LevelCount"));
                                    this.start_level = reader.GetAttribute("start_level");
                                    this.end_level = reader.GetAttribute("end_level");
                                    break;
                                }
                            case "Level":
                                {
                                    if (this.mark != wallMark) break;
                                    this.Level = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                            case "bt":
                                {
                                    if (this.mark != wallMark) break;
                                    this.bt = ReadMultiAttributeDouble(reader, this.LevelCount);
                                    break;
                                }
                            case "bd":
                                {
                                    if (this.mark != wallMark) break;
                                    this.bd = ReadMultiAttributeDouble(reader, this.LevelCount);
                                    break;
                                }
                            case "D":
                                {
                                    if (this.mark != wallMark) break;
                                    this.D = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                            case "number_A":
                                {
                                    if (this.mark != wallMark) break;
                                    this.number_A = ReadMultiAttributeInt(reader, this.LevelCount);
                                    break;
                                }
                            case "number_B":
                                {
                                    if (this.mark != wallMark) break;
                                    this.number_B = ReadMultiAttributeInt(reader, this.LevelCount);
                                    break;
                                }
                            case "SCDaiNgoai":
                                {
                                    if (this.mark != wallMark) break;
                                    this.SCDaiNgoai = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                            case "DKDaiNgoai":
                                {
                                    if (this.mark != wallMark) break;
                                    this.DKDaiNgoai = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                            case "SCDaiBien":
                                {
                                    if (this.mark != wallMark) break;
                                    this.SCDaiBien = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                            case "DKDaiBien":
                                {
                                    if (this.mark != wallMark) break;
                                    this.DKDaiBien = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                            case "SCDaiCon":
                                {
                                    if (this.mark != wallMark) break;
                                    this.SCDaiCon = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                            case "DKDaiCon":
                                {
                                    if (this.mark != wallMark) break;
                                    this.DKDaiCon = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                            case "spacing1_con":
                                {
                                    if (this.mark != wallMark) break;
                                    this.spacing1_con = ReadMultiAttributeDouble(reader, this.LevelCount);
                                    break;
                                }
                            case "spacing2_con":
                                {
                                    if (this.mark != wallMark) break;
                                    this.spacing2_con = ReadMultiAttributeDouble(reader, this.LevelCount);
                                    break;
                                }
                            case "spacing3_bao":
                                {
                                    if (this.mark != wallMark) break;
                                    this.spacing3_bao = ReadMultiAttributeDouble(reader, this.LevelCount);
                                    break;
                                }
                            case "spacing4_bao":
                                {
                                    if (this.mark != wallMark) break;
                                    this.spacing4_bao = ReadMultiAttributeDouble(reader, this.LevelCount);
                                    break;
                                }
                            case "spacing5_bien":
                                {
                                    if (this.mark != wallMark) break;
                                    this.spacing5_bien = ReadMultiAttributeDouble(reader, this.LevelCount);
                                    break;
                                }
                            case "spacing6_bien":
                                {
                                    if (this.mark != wallMark) break;
                                    this.spacing6_bien = ReadMultiAttributeDouble(reader, this.LevelCount);
                                    break;
                                }
                            case "daiNgang":
                                {
                                    if (this.mark != wallMark) break;
                                    this.daiNgang = ReadMultiAttributeString(reader, this.LevelCount);
                                    break;
                                }
                        }
                        break;
                }
            }
            L1:
            reader.Close();
        }
        public List<string> ReadMultiAttributeString(XmlReader reader, int count)
        {
            List<string> ss = new List<string>();
            for (int i = 0; i < count; i++)
            {
                ss.Add(reader.GetAttribute("ID" + (i + 1).ToString()));
            }
            return ss;
        }
        public List<int> ReadMultiAttributeInt(XmlReader reader, int count)
        {
            List<int> ss = new List<int>();
            for (int i = 0; i < count; i++)
            {
                ss.Add(int.Parse(reader.GetAttribute("ID" + (i + 1).ToString())));
            }
            return ss;
        }
        public List<double> ReadMultiAttributeDouble(XmlReader reader, int count)
        {
            List<double> ss = new List<double>();
            for (int i = 0; i < count; i++)
            {
                ss.Add(double.Parse(reader.GetAttribute("ID" + (i + 1).ToString())));
            }
            return ss;
        }
    }
    public class DataProcess
    {
        #region "khai báo biến toàn cục"
        UIApplication uiapp;
        UIDocument uidoc;
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        Selection sel;
        #endregion
        #region "khai báo đặc tính"
        public string mark { get; set; }
        public List<RebarBarType> type_D { get; set; }
        public List<double> D { get; set; }

        public List<RebarShape> SCDaiNgoai { get; set; }
        public List<RebarBarType> type_DKDaiNgoai { get; set; }
        public List<double> DKDaiNgoai { get; set; }

        public List<RebarShape> SCDaiBien { get; set; }
        public List<RebarBarType> type_DKDaiBien { get; set; }
        public List<double> DKDaiBien { get; set; }

        public List<RebarShape> SCDaiCon { get; set; }
        public List<RebarBarType> type_DKDaiCon { get; set; }
        public List<double> DKDaiCon { get; set; }

        public List<double> hVach { get; set; }
        public List<double> width { get; set; }
        public List<double> length { get; set; }
        public List<double> cover { get; set; }
        public List<double> a { get; set; }
        public List<double> A { get; set; }
        public List<double> B { get; set; }
        public List<double> C { get; set; }

        public List<Autodesk.Revit.DB.Line> driving { get; set; }
        public List<XYZ> vecX { get; set; }
        public List<XYZ> vecY { get; set; }

        public List<XYZ> cp0 { get; set; }
        public List<XYZ> bp0 { get; set; }
        public List<XYZ> mp01 { get; set; }
        public List<XYZ> mp02 { get; set; }
        public List<XYZ> cp1 { get; set; }
        public List<XYZ> bp1 { get; set; }
        public List<XYZ> mp11 { get; set; }
        public List<XYZ> mp12 { get; set; }

        #endregion
        public DataProcess(ImportExcel excel, ExternalCommandData commandData, List<Element> column)
        {
            #region "khai báo List"
            mark = excel.mark;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            app = uiapp.Application;
            doc = uidoc.Document;
            sel = uidoc.Selection;
            //-----------------------------------------------------------------------------------------------------
            type_D = new List<RebarBarType>();
            D = new List<double>();

            SCDaiNgoai = new List<RebarShape>();
            type_DKDaiNgoai = new List<RebarBarType>();
            DKDaiNgoai = new List<double>();

            SCDaiBien = new List<RebarShape>();
            type_DKDaiBien = new List<RebarBarType>();
            DKDaiBien = new List<double>();

            SCDaiCon = new List<RebarShape>();
            type_DKDaiCon = new List<RebarBarType>();
            DKDaiCon = new List<double>();
            //-----------------------------------------------------------------------------------------------------
            hVach = new List<double>();
            width = new List<double>();
            length = new List<double>();
            cover = new List<double>();
            a = new List<double>();

            A = new List<double>();
            B = new List<double>();

            driving = new List<Autodesk.Revit.DB.Line>();
            vecX = new List<XYZ>();
            vecY = new List<XYZ>();

            cp0 = new List<XYZ>();
            bp0 = new List<XYZ>();
            mp01 = new List<XYZ>();
            mp02 = new List<XYZ>();
            cp1 = new List<XYZ>();
            bp1 = new List<XYZ>();
            mp11 = new List<XYZ>();
            mp12 = new List<XYZ>();
            FilteredElementCollector colBarType = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType));
            FilteredElementCollector colBarShape = new FilteredElementCollector(doc).OfClass(typeof(RebarShape));
            #endregion
            for (int i = 0; i < column.Count; i++)
            {
                type_D.Add(colBarType.Where(x => x.Name == excel.D[i]).Cast<RebarBarType>().First());
                D.Add(type_D[i].LookupParameter("Bar Diameter").AsDouble());

                SCDaiNgoai.Add(colBarShape.Where(x => x.Name == excel.SCDaiNgoai[i]).Cast<RebarShape>().First());
                type_DKDaiNgoai.Add(colBarType.Where(x => x.Name == excel.DKDaiNgoai[i]).Cast<RebarBarType>().First());
                DKDaiNgoai.Add(type_DKDaiNgoai[i].LookupParameter("Bar Diameter").AsDouble());

                SCDaiBien.Add(colBarShape.Where(x => x.Name == excel.SCDaiBien[i]).Cast<RebarShape>().First());
                type_DKDaiBien.Add(colBarType.Where(x => x.Name == excel.DKDaiBien[i]).Cast<RebarBarType>().First());
                DKDaiBien.Add(type_DKDaiBien[i].LookupParameter("Bar Diameter").AsDouble());

                SCDaiCon.Add(colBarShape.Where(x => x.Name == excel.SCDaiCon[i]).Cast<RebarShape>().First());
                type_DKDaiCon.Add(colBarType.Where(x => x.Name == excel.DKDaiCon[i]).Cast<RebarBarType>().First());
                DKDaiCon.Add(type_DKDaiCon[i].LookupParameter("Bar Diameter").AsDouble());

                hVach.Add(column[i].LookupParameter("Unconnected Height").AsDouble());
                ElementType eType = doc.GetElement(column[i].GetTypeId()) as ElementType;
                width.Add(eType.LookupParameter("Width").AsDouble());
                length.Add(column[i].LookupParameter("Length").AsDouble());
                cover.Add(rebarCover(column[i]));
                a.Add(cover[i] + D[i] / 2 + DKDaiNgoai[i]);

                A.Add(length[i] / (double)excel.tl - a[i]);
                B.Add(length[i] * (1 - 2 / (double)excel.tl));
                C.Add(width[i] - 2 * a[i]);

                driving.Add(drivingLine(column[i]));
                vecX.Add(CheckGeometry.GetDirection(driving[i]));
                vecY.Add(XYZ.BasisZ.CrossProduct(vecX[i]));

                cp0.Add(GeomUtil.OffsetPoint(driving[i].GetEndPoint(0), vecY[i], C[i] / 2));
                bp0.Add(GeomUtil.OffsetPoint(cp0[i], vecX[i], a[i]));
                mp01.Add(GeomUtil.OffsetPoint(bp0[i], vecX[i], A[i]));
                mp02.Add(GeomUtil.OffsetPoint(mp01[i], vecX[i], B[i]));
                cp1.Add(GeomUtil.OffsetPoint(driving[i].GetEndPoint(1), -vecY[i], C[i] / 2));
                bp1.Add(GeomUtil.OffsetPoint(cp1[i], -vecX[i], a[i]));
                mp11.Add(GeomUtil.OffsetPoint(bp1[i], -vecX[i], A[i]));
                mp12.Add(GeomUtil.OffsetPoint(mp11[i], -vecX[i], B[i]));
            }
        }
        public double rebarCover(Element e)
        {
            Autodesk.Revit.DB.Parameter rcPara = e.LookupParameter("Rebar Cover - Exterior Face");
            ElementId rcId = rcPara.AsElementId();
            RebarCoverType rcType = doc.GetElement(rcId) as ElementType as RebarCoverType;
            return rcType.CoverDistance; ;
        }
        public Autodesk.Revit.DB.Line drivingLine(Element wall)
        {
            LocationCurve lc = wall.Location as LocationCurve; Curve c = lc.Curve;
            Autodesk.Revit.DB.Parameter boP = wall.LookupParameter("Base Offset");
            c = GeomUtil.OffsetCurve(c, XYZ.BasisZ, boP.AsDouble());
            List<XYZ> ps = new List<XYZ> { c.GetEndPoint(0), c.GetEndPoint(1) };
            ps.Sort(new ZYXComparer());
            return Autodesk.Revit.DB.Line.CreateBound(ps[0], ps[1]);
        }
    }
}
