using System;
using System.IO;
using System.Data;
using System.Web;
using System.Web.UI.WebControls;
using System.Runtime.InteropServices;
using Oracle.DataAccess.Client;
using Microsoft.Reporting.WebForms;
using SRL.UserControls;
//using System.Net.Mail;
//using PQScan.PDFToImage;
using System.Drawing;
using System.Drawing.Imaging;
//using System.Configuration;

namespace ERPAppAddition.ERPAddition
{
    public class GF
    {
        //private void ReportCreator(System.Data.DataSet dsSource, string strSQL, ReportViewer rvCtrl, string strRptName, string strRptDatSrcName)
        //{
        //    gOraCon.Open();
        //    gOraCmd = gOraCon.CreateCommand();
        //    gOraCmd.CommandType = CommandType.Text;
        //    rvCtrl.Reset(); //레포트뷰어를 초기화한다.

        //    try
        //    {
        //        gOraCmd.CommandText = strSQL;
        //        gOraDR = gOraCmd.ExecuteReader();
        //        dsSource.Tables[0].Load(gOraDR);
        //        gOraDR.Close();

        //        rvCtrl.LocalReport.ReportEmbeddedResource = GC.reportPath + strRptName;
        //        rvCtrl.LocalReport.DisplayName = "REPORT_" + this.Title + DateTime.Now.ToShortDateString();

        //        ReportDataSource rds = new ReportDataSource();
        //        rds.Name = strRptDatSrcName;
        //        rds.Value = dsSource.Tables[0];
        //        rvCtrl.LocalReport.DataSources.Add(rds);
        //        rvCtrl.LocalReport.Refresh();
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("{0} Second exception caught.", e);
        //    }
        //    finally
        //    {
        //        if (gOraCon.State == ConnectionState.Open)
        //            gOraCon.Close();
        //    }
        //}
        public static string GetDrpDwnChkLstValue(CheckBoxList cblData, TextBox txtData)    //선택된 체크박스리스트의 아이템의 value값을 읽어온다.
        {
            string str = string.Empty;
            foreach (ListItem li in cblData.Items)
            {
                if (li.Selected)
                {
                    if (str == "")
                    {
                        str = "'" + li.Value + "'";
                        txtData.Text = li.Value;
                    }
                    else
                    {
                        str += ",'" + li.Value + "'";
                        txtData.Text += "," + li.Value;
                    }
                }
            }
            return str;
        }
        public static void SetHiddenField(OracleConnection oraCon, string strSQL, string strVField, HiddenField hfValue)
        {
            GV.gOraCmd = new OracleCommand(strSQL, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();

            hfValue.Value = "";
            while (GV.gOraDR.Read())
            {
                hfValue.Value = hfValue.Value + "'" + GV.gOraDR[strVField].ToString() + "', ";

            }
            if (hfValue.Value != "") hfValue.Value = hfValue.Value.Substring(0, hfValue.Value.Length - 2);
        }
        public static string Replace(string strColVal, string strDefSrc, string strDefTgt)
        {
            if (strColVal == null || strColVal.Trim() == "") return "";

            string strTemp = strColVal;
            return strTemp.Replace(strDefSrc, strDefTgt);
        }
        public static string AddDays(string strDT, int iDays, GC.DateTimeFmt dtfValue = GC.DateTimeFmt.yyyyMMddhhmmss)
        {
            string strRet = "";

            switch (dtfValue)
            {
                case GC.DateTimeFmt.yyyyMMddhhmmss:
                    strRet = Convert.ToDateTime(strDT.Substring(0, 4) + "/" + strDT.Substring(4, 2) + "/" + strDT.Substring(6, 2) + " "
                                              + strDT.Substring(8, 2) + ":" + strDT.Substring(10, 2) + ":" + strDT.Substring(12, 2)).AddDays(iDays).ToString("yyyyMMddhhmmss");
                    break;
                case GC.DateTimeFmt.yyyyMMdd:
                    strRet = Convert.ToDateTime(strDT.Substring(0, 4) + "/" + strDT.Substring(4, 2) + "/" + strDT.Substring(6, 2)).AddDays(iDays).ToString("yyyyMMdd");
                    break;
                case GC.DateTimeFmt.yyyyMMddhhmm:
                    strRet = Convert.ToDateTime(strDT.Substring(0, 4) + "/" + strDT.Substring(4, 2) + "/" + strDT.Substring(6, 2) + " "
                                              + strDT.Substring(8, 2) + ":" + strDT.Substring(10, 2)).AddDays(iDays).ToString("yyyyMMddhhmm");
                    break;
                case GC.DateTimeFmt.yyyyMMddhh:
                    strRet = Convert.ToDateTime(strDT.Substring(0, 4) + "/" + strDT.Substring(4, 2) + "/" + strDT.Substring(6, 2) + " "
                                              + strDT.Substring(8, 2)).AddDays(iDays).ToString("yyyyMMddhh");
                    break;
            }

            return strRet;
        }

        #region "Functions_For_Display"
        public static void CreateReport(System.Data.DataSet dsSource, short sTableNo, string strSQL, ReportViewer rvCtrl,
                                    string strRDLC, string strRptDatSrcName, string strDispName)
        {
            GV.gOraCon.Open();
            GV.gOraCmd = GV.gOraCon.CreateCommand();
            GV.gOraCmd.CommandType = CommandType.Text;
            //rvCtrl.Reset(); //레포트뷰어를 초기화한다.

            try
            {
                GV.gOraCmd.CommandText = strSQL;
                //GV.gOraCmd.CommandTimeout = 360;
                GV.gOraDR = GV.gOraCmd.ExecuteReader();
                dsSource.Tables[sTableNo].Load(GV.gOraDR);
                GV.gOraDR.Close();

                rvCtrl.ProcessingMode = ProcessingMode.Local;
                //rvCtrl.LocalReport.ReportPath = Server.MapPath(GC.reportPath + strRptName);
                rvCtrl.LocalReport.ReportEmbeddedResource = GC.reportPath + strRDLC;
                rvCtrl.LocalReport.DisplayName = strDispName;

                ReportDataSource rds = new ReportDataSource(strRptDatSrcName, dsSource.Tables[sTableNo]);
                //rvCtrl.WaitControlDisplayAfter = 36000;
                rvCtrl.LocalReport.DataSources.Clear();
                rvCtrl.LocalReport.DataSources.Add(rds);
                //rvCtrl.LocalReport.Refresh();            

            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Second exception caught.", e);
            }
            finally
            {
                if (GV.gOraCon.State == ConnectionState.Open) GV.gOraCon.Close();
            }
        }

        // 1개의 레포트뷰어에 2개의 DateSet을 연결하기 위한 함수 생성 (20150730, 안창주)
        public static void CreateReport_DataSet2(System.Data.DataSet dsSource, short sTableNo, short sTableNo2, string strSQL, string strSQL_Prod, ReportViewer rvCtrl,
                            string strRDLC, string strRptDatSrcName, string strRptDatSrcName2, string strDispName)
        {
            //GV.gOraCon.Open();
            //GV.gOraCmd = GV.gOraCon.CreateCommand();
            //GV.gOraCmd.CommandType = CommandType.Text;
            //rvCtrl.Reset(); //레포트뷰어를 초기화한다..

            try
            {
                //도급비데이터를 가져오기 위해 디스플레이 DB 오픈
                GV.gOraCon.Open();
                GV.gOraCmd = GV.gOraCon.CreateCommand();
                GV.gOraCmd.CommandType = CommandType.Text;

                GV.gOraCmd.CommandText = strSQL;
                GV.gOraDR = GV.gOraCmd.ExecuteReader();
                dsSource.Tables[sTableNo].Load(GV.gOraDR);
                GV.gOraDR.Close();
                GV.gOraCon.Close();


                //실적데이터를 가져오기 위해 RPT DB 오픈
                GV.gOraCo2.Open();
                GV.gOraCmd = GV.gOraCo2.CreateCommand();
                GV.gOraCmd.CommandType = CommandType.Text;

                GV.gOraCmd.CommandText = strSQL_Prod;
                GV.gOraDR = GV.gOraCmd.ExecuteReader();
                dsSource.Tables[sTableNo2].Load(GV.gOraDR);
                GV.gOraDR.Close();
                GV.gOraCo2.Close();


                rvCtrl.ProcessingMode = ProcessingMode.Local;
                //rvCtrl.LocalReport.ReportPath = Server.MapPath(GC.reportPath + strRptName);
                rvCtrl.LocalReport.ReportEmbeddedResource = GC.reportPath + strRDLC;
                rvCtrl.LocalReport.DisplayName = strDispName;

                ReportDataSource rds = new ReportDataSource(strRptDatSrcName, dsSource.Tables[sTableNo]);
                ReportDataSource rds2 = new ReportDataSource(strRptDatSrcName2, dsSource.Tables[sTableNo2]);
                //rvCtrl.WaitControlDisplayAfter = 36000;
                rvCtrl.LocalReport.DataSources.Clear();
                rvCtrl.LocalReport.DataSources.Add(rds);
                rvCtrl.LocalReport.DataSources.Add(rds2);
                //rvCtrl.LocalReport.Refresh();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Second exception caught.", e);
            }
            //finally
            //{
            //    if (GV.gOraCon.State == ConnectionState.Open) GV.gOraCon.Close();
            //}
        }
        public static void SetMatID2Ctrl(OracleConnection oraCon, MultiCheckCombo mccMatID, GC.MatType mtValue = GC.MatType.COMP)
        {
            string strSqlMatID = "";

            switch (mtValue)
            {
                case GC.MatType.COMP:
                    strSqlMatID = GC.SQL_MAT_ID_COMP;
                    break;
                case GC.MatType.MTRL:
                    strSqlMatID = GC.SQL_MAT_ID_MTRL;
                    break;
                case GC.MatType.SHET:
                    strSqlMatID = GC.SQL_MAT_ID_SHET;
                    break;
            }
            GV.gOraCmd = new OracleCommand(strSqlMatID, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccMatID.AddItems(GV.gOraDR, "MAT_ID", "MAT_ID");
        }
        public static void SetMatID2Ctrl(OracleConnection oraCon, DropDownList ddlMatID, GC.MatType mtValue = GC.MatType.COMP)
        {
            string strSqlMatID = "";

            switch (mtValue)
            {
                case GC.MatType.COMP:
                    strSqlMatID = GC.SQL_MAT_ID_COMP;
                    break;
                case GC.MatType.MTRL:
                    strSqlMatID = GC.SQL_MAT_ID_MTRL;
                    break;
                case GC.MatType.SHET:
                    strSqlMatID = GC.SQL_MAT_ID_SHET;
                    break;
            }
            GV.gOraCmd = new OracleCommand(strSqlMatID, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            ddlMatID.DataSource = GV.gOraDR;
            ddlMatID.DataValueField = ddlMatID.DataTextField = "MAT_ID";
            ddlMatID.DataBind();
        }
        public static void SetFlow2Ctrl(OracleConnection oraCon, MultiCheckCombo mccFlow)
        {
            GV.gOraCmd = new OracleCommand(GC.SQL_SFLOW, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccFlow.AddItems(GV.gOraDR, "FLOW_DESC", "FLOW");
        }
        public static void SetOper2Ctrl(OracleConnection oraCon, MultiCheckCombo mccOper, GC.MatType mtValue = GC.MatType.COMP)
        {
            GV.gOraCmd = new OracleCommand((mtValue == GC.MatType.COMP) ? GC.SQL_SOPER_COMP : GC.SQL_SOPER_MTRL, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccOper.AddItems(GV.gOraDR, "OPER_DESC", "OPER");
        }
        public static void SetOper2Ctrl(OracleConnection oraCon, DropDownList ddlOper, GC.MatType mtValue = GC.MatType.COMP)
        {
            GV.gOraCmd = new OracleCommand((mtValue == GC.MatType.COMP) ? GC.SQL_SOPER_COMP : GC.SQL_SOPER_MTRL, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            ddlOper.DataSource = GV.gOraDR;
            ddlOper.DataValueField = "OPER";
            ddlOper.DataTextField = "OPER_DESC";
            ddlOper.DataBind();
        }
        public static void SetOper2Ctrl(OracleConnection oraCon, DropDownList ddlOper, string strInitOper, GC.MatType mtValue = GC.MatType.COMP)
        {
            GV.gOraCmd = new OracleCommand((mtValue == GC.MatType.COMP) ? GC.SQL_SOPER_COMP : GC.SQL_SOPER_MTRL, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();

            int i = 0;
            while (GV.gOraDR.Read())
            {
                ddlOper.Items.Add(GV.gOraDR["OPER_DESC"].ToString());
                ddlOper.Items[i].Value = GV.gOraDR["OPER"].ToString();
                i++;
            }
            if (strInitOper != "") if (!ddlOper.Page.IsPostBack) ddlOper.SelectedValue = ddlOper.Items.FindByValue(strInitOper).Value;
        }
        public static void SetLoss2Ctrl(OracleConnection oraCon, MultiCheckCombo mccOper, GC.MatType mtValue = GC.MatType.COMP)
        {
            GV.gOraCmd = new OracleCommand(GC.SQL_SLOSS_ALL, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccOper.AddItems(GV.gOraDR, "LOSS_DESC", "LOSS_CODE");
        }
        public static void SetLoss2Ctrl(OracleConnection oraCon, DropDownList ddlOper, GC.MatType mtValue = GC.MatType.COMP)
        {
            GV.gOraCmd = new OracleCommand(GC.SQL_SLOSS_ALL, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            ddlOper.DataSource = GV.gOraDR;
            ddlOper.DataValueField = "LOSS_CODE";
            ddlOper.DataTextField = "LOSS_DESC";
            ddlOper.DataBind();
        }
        public static void SetCreateCode2Ctrl(OracleConnection oraCon, MultiCheckCombo mccOper, GC.MatType mtValue = GC.MatType.COMP)
        {
            GV.gOraCmd = new OracleCommand(GC.SQL_SCREATE_CODE, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccOper.AddItems(GV.gOraDR, "CC_DESC", "CREATE_CODE", mccOper.InitTextList);
        }
        #endregion

        #region "Functions_For_HRMS"
        public static void SetAreaIDCtrl(OracleConnection oraCon, MultiCheckCombo mcc_dr_Area, string Depart)  // Daily Report 조회조건 관련 추가 (2015.06.17_안창주)
        {
            string strSqlAreaID = "";

            if (Depart == "Semi")
            {
                strSqlAreaID = GC.SQL_AREA_CODE_SEMI;
            }
            else
            {
                strSqlAreaID = GC.SQL_AREA_CODE_DISP;
            }

            GV.gOraCmd = new OracleCommand(strSqlAreaID, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mcc_dr_Area.AddItems(GV.gOraDR, "AREA", "AREA");
        }
        public static void SetOprGroupIDCtrl(OracleConnection oraCon, MultiCheckCombo mcc_dr_Oprgrp, string Depart)  // Daily Report 조회조건 관련 추가 (2015.06.17_안창주)
        {
            string strSqlOprID = "";

            if (Depart == "Semi")
            {
                strSqlOprID = GC.SQL_OPERGRP_CODE_SEMI;
            }
            else
            {
                strSqlOprID = GC.SQL_OPERGRP_CODE_DISP;
            }

            GV.gOraCmd = new OracleCommand(strSqlOprID, oraCon);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mcc_dr_Oprgrp.AddItems(GV.gOraDR, "OPER_GROUP", "OPER_GROUP");
        }
        public static void SetPartIDCtrl(OracleConnection oraCon, MultiCheckCombo mcc_dr_Part, string Depart)  // Daily Report 조회조건 관련 추가 (2015.06.17_안창주)
        {
            string strSqlPartID = "";

            if (Depart == "Semi")
            {
                strSqlPartID = GC.SQL_PART_CODE_SEMI;

                GV.gOraCmd = new OracleCommand(strSqlPartID, oraCon);
                GV.gOraDR = GV.gOraCmd.ExecuteReader();
                mcc_dr_Part.AddItems(GV.gOraDR, "PART", "PART");
            }
        }
        public static void SetCostIDCtrl(OracleConnection oraCon, MultiCheckCombo mcc_dr_Cost, string Depart)  // Daily Report 조회조건 관련 추가 (2015.08.16_안창주)
        {
            string strSqlPartID = "";

            if (Depart == "Semi")
            {
                strSqlPartID = GC.SQL_COST_CODE_SEMI;

                GV.gOraCmd = new OracleCommand(strSqlPartID, oraCon);
                GV.gOraDR = GV.gOraCmd.ExecuteReader();
                //mcc_dr_Cost.AddItems(GV.gOraDR, "S_GROUP", "SUB_CODE");
                mcc_dr_Cost.AddItems(GV.gOraDR, "S_GROUP", "S_GROUP");
            }
        }
        public static void SetDepartIDCtrl(OracleConnection oraCon, MultiCheckCombo mcc_dr_Depart, string Depart)  // Daily Report 조회조건 관련 추가 (2015.08.16_안창주)
        {
            string strSqlPartID = "";

            if (Depart == "Semi")
            {
                strSqlPartID = GC.SQL_DEPART_CODE_SEMI;

                GV.gOraCmd = new OracleCommand(strSqlPartID, oraCon);
                GV.gOraDR = GV.gOraCmd.ExecuteReader();
                //mcc_dr_Depart.AddItems(GV.gOraDR, "S_GROUP", "SUB_CODE");
                mcc_dr_Depart.AddItems(GV.gOraDR, "S_GROUP", "S_GROUP");
            }
        }
        #endregion

        #region "Functions_For_Semiconductor"
        public static void CreateReport2(System.Data.DataSet dsSource, short sTableNo, string strSQL, ReportViewer rvCtrl,
                                    string strRDLC, string strRptDatSrcName, string strDispName)
        {
            GV.gOraCo2.Open();
            GV.gOraCmd = GV.gOraCo2.CreateCommand();
            GV.gOraCmd.CommandType = CommandType.Text;
            //rvCtrl.Reset(); //레포트뷰어를 초기화한다.

            try
            {
                GV.gOraCmd.CommandText = strSQL;
                //GV.gOraCmd.CommandTimeout = 360;
                GV.gOraDR = GV.gOraCmd.ExecuteReader();
                dsSource.Tables[sTableNo].Load(GV.gOraDR);
                GV.gOraDR.Close();

                rvCtrl.ProcessingMode = ProcessingMode.Local;
                //rvCtrl.LocalReport.ReportPath = Server.MapPath(GC.reportPath + strRptName);
                rvCtrl.LocalReport.ReportEmbeddedResource = GC.reportPath + strRDLC;
                rvCtrl.LocalReport.DisplayName = strDispName;

                ReportDataSource rds = new ReportDataSource(strRptDatSrcName, dsSource.Tables[sTableNo]);
                //rvCtrl.WaitControlDisplayAfter = 36000;
                rvCtrl.LocalReport.DataSources.Clear();
                rvCtrl.LocalReport.DataSources.Add(rds);
                //rvCtrl.LocalReport.Refresh();

                //EmailReport(rds);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Second exception caught.", e);
            }
            finally
            {
                if (GV.gOraCo2.State == ConnectionState.Open) GV.gOraCo2.Close();
            }
        }
        public static void SetPartID2Ctrl(OracleConnection oraCo2, MultiCheckCombo mccPartID, GC.PlantType ptValue, GC.MethodType mtValue = GC.MethodType.CRNT)
        {
            string strSqlPartID = "";

            switch (ptValue)
            {
                case GC.PlantType.P01:
                    strSqlPartID = GC.SQL_PART_ID_CRNT_P01;
                    break;
                case GC.PlantType.P02:
                    strSqlPartID = GC.SQL_PART_ID_CRNT_P02;
                    break;
                case GC.PlantType.P09:
                    strSqlPartID = GC.SQL_PART_ID_CRNT_P09;
                    break;
                case GC.PlantType.P12:
                    strSqlPartID = GC.SQL_PART_ID_CRNT_P12;
                    break;
            }
            GV.gOraCmd = new OracleCommand(strSqlPartID, oraCo2);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccPartID.AddItems(GV.gOraDR, "PART_ID", "PART_ID");
        }
        public static void SetPartID2CtrlEx(OracleConnection oraCo2, MultiCheckCombo mccPartID, GC.ProdType ptValue, GC.MethodType mtValue = GC.MethodType.CRNT)
        {
            string strSqlPartID = "";

            switch (ptValue)
            {
                case GC.ProdType.DDI:
                    strSqlPartID = GC.SQL_PART_ID_CRNT_DDI;
                    break;
                case GC.ProdType.WLP:
                    strSqlPartID = GC.SQL_PART_ID_CRNT_WLP;
                    break;
                case GC.ProdType.FOWLP:
                    strSqlPartID = GC.SQL_PART_ID_CRNT_FOWLP;
                    break;
            }
            GV.gOraCmd = new OracleCommand(strSqlPartID, oraCo2);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccPartID.AddItems(GV.gOraDR, "PART_ID", "PART_ID");
        }
        public static void SetOper2Ctrl(OracleConnection oraCo2, MultiCheckCombo mccOper, GC.PlantType ptValue, GC.MethodType mtValue = GC.MethodType.CRNT)
        {
            string strSqlOper = "";

            switch (ptValue)
            {
                case GC.PlantType.P01:
                    strSqlOper = GC.SQL_SOPER_CRNT_P01;
                    break;
                case GC.PlantType.P02:
                    strSqlOper = GC.SQL_SOPER_CRNT_P02;
                    break;
                case GC.PlantType.P09:
                    strSqlOper = GC.SQL_SOPER_CRNT_P09;
                    break;
                case GC.PlantType.P12:
                    strSqlOper = GC.SQL_SOPER_CRNT_P12;
                    break;
            }
            GV.gOraCmd = new OracleCommand(strSqlOper, oraCo2);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccOper.AddItems(GV.gOraDR, "OPER_DESC", "OPER");
        }
        public static void SetOper2CtrlEx(OracleConnection oraCo2, MultiCheckCombo mccOper, GC.ProdType ptValue, GC.MethodType mtValue = GC.MethodType.CRNT)
        {
            string strSqlOper = "";

            switch (ptValue)
            {
                case GC.ProdType.DDI:
                    strSqlOper = GC.SQL_SOPER_CRNT_DDI;
                    break;
                case GC.ProdType.WLP:
                    strSqlOper = GC.SQL_SOPER_CRNT_WLP;
                    break;
                case GC.ProdType.FOWLP:
                    strSqlOper = GC.SQL_SOPER_CRNT_FOWLP;
                    break;
            }
            GV.gOraCmd = new OracleCommand(strSqlOper, oraCo2);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccOper.AddItems(GV.gOraDR, "OPER_DESC", "OPER");
        }
        public static void SetCreateCode2Ctrl(OracleConnection oraCo2, MultiCheckCombo mccOper, GC.PlantType ptValue, GC.MethodType mtValue = GC.MethodType.CRNT)
        {
            string strSqlCC = "";

            switch (ptValue)
            {
                case GC.PlantType.P01:
                    strSqlCC = GC.SQL_CREATE_CODE_CRNT_P01;
                    break;
                case GC.PlantType.P02:
                    strSqlCC = GC.SQL_CREATE_CODE_CRNT_P02;
                    break;
                case GC.PlantType.P09:
                    strSqlCC = GC.SQL_CREATE_CODE_CRNT_P09;
                    break;
                case GC.PlantType.P12:
                    strSqlCC = GC.SQL_CREATE_CODE_CRNT_P12;
                    break;
            }
            GV.gOraCmd = new OracleCommand(strSqlCC, oraCo2);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccOper.AddItems(GV.gOraDR, "CREATE_CODE", "CREATE_CODE");
        }
        public static void SetCreateCode2CtrlEx(OracleConnection oraCo2, MultiCheckCombo mccOper, GC.ProdType ptValue, GC.MethodType mtValue = GC.MethodType.CRNT)
        {
            string strSqlCC = "";

            switch (ptValue)
            {
                case GC.ProdType.DDI:
                    strSqlCC = GC.SQL_CREATE_CODE_CRNT_DDI;
                    break;
                case GC.ProdType.WLP:
                    strSqlCC = GC.SQL_CREATE_CODE_CRNT_WLP;
                    break;
                case GC.ProdType.FOWLP:
                    strSqlCC = GC.SQL_CREATE_CODE_CRNT_FOWLP;
                    break;
            }
            GV.gOraCmd = new OracleCommand(strSqlCC, oraCo2);
            GV.gOraDR = GV.gOraCmd.ExecuteReader();
            mccOper.AddItems(GV.gOraDR, "CREATE_CODE", "CREATE_CODE");
        }
        public static void EmailReport(ReportDataSource rds)
        {
            LocalReport lr = new LocalReport();
            lr.ReportPath = HttpContext.Current.Server.MapPath("~/Reports/WM/DailyWIPTrend.rdlc");
            lr.DataSources.Add(rds);

            Warning[] warnings;
            string[] streamids;
            string mimeType, encoding, extension;

            //http://<Server Name>/reportserver?/SampleReports/Sales Order Detail&rs:Command=Render&rs:Format=HTML4.0&rc:Toolbar=False

            //PageCountMode pcm = PageCountMode.Estimate; PDF
            byte[] bytes = lr.Render("IMAGE",
                @"<DeviceInfo>
                    <PageWidth>30cm</PageWidth>
                    <PageHeight>30cm</PageHeight>
                    <MarginTop>0in</MarginTop>
                    <MarginBottom>0in</MarginBottom>
                    <MarginLeft>0in</MarginLeft>
                    <MarginRight>0in</MarginRight>
                    <OutputFormat>JPEG</OutputFormat>
                </DeviceInfo>",
                out mimeType, out encoding, out extension, out streamids, out warnings);
            File.WriteAllBytes("D:\\myReport.jpg", bytes);

            //MemoryStream s = new MemoryStream(bytes);
            //s.Seek(0, SeekOrigin.Begin);
            //Attachment b;//, a = new Attachment(s, "myReport.pdf");

            //PDFDocument pdfDoc = new PDFDocument();

            //// Load PDF document from local file.
            //pdfDoc.LoadPDF("D:\\myReport.pdf");
            //pdfDoc.ToMultiPageTiff("D:\\output.tif");

            //// Get the total page count.
            //int count = pdfDoc.PageCount;

            //for (int i = 0; i < count; i++)
            //{
            //    // Convert PDF page to image.
            //    Bitmap jpgImage = pdfDoc.ToImage(i);

            //    // Save image with jpg file type.
            //    jpgImage.Save("output" + i + ".jpg", ImageFormat.Jpeg);
            //}
            //b = new Attachment("output0.jpg");

            //MailMessage message = new MailMessage("kennethlee@nepes.co.kr", "choiyd@nepes.co.kr", "A report for you!", "Here is a report for you");
            //message.Attachments.Add(b);
            //SmtpClient client = new SmtpClient();
            //client.Send(message);
        }
        #endregion
    }
}
