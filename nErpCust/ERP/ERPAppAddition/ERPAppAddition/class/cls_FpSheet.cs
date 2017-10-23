using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using FarPoint.Web.Spread;


	public class cls_FpSheet
	{
        public void AddColumnHeader(FarPoint.Web.Spread.SheetView sv, string sDataFild, string sHeaderNM, int Width, HorizontalAlign hAlign = HorizontalAlign.Left,  bool Lock = false, int row = 0, FarPoint.Web.Spread.ICellType CellType = null, bool Visible = true, int headerColSpan = 0, int headerRowSpan =0)
        {
            DataTable dt = new DataTable();

            if (sv.DataSource == null)
            {
                dt = new DataTable();
                //sv.AutoGenerateColumns = false;

               
                sv.AlternatingRows.Count = 2;
                sv.AlternatingRows[0].BackColor = Color.Snow;
                sv.AlternatingRows[1].BackColor = Color.Lavender;

                sv.SelectionBackColor = Color.LightBlue;

                sv.PageSize = 25;

                sv.FpSpread.DataBind();

            }
            else
            {
                dt = (DataTable)sv.DataSource;

            }

            dt.Columns.Add(sDataFild);


            if (CellType == null)
            {
                FarPoint.Web.Spread.TextCellType txtCellType = new FarPoint.Web.Spread.TextCellType();

                CellType = txtCellType;
            }
            else if (CellType.ToString() == "DateCalendarCellType")
            {
                dt.Columns[dt.Columns.Count - 1].DataType = typeof(DateTime);
            }


            FarPoint.Web.Spread.Model.DefaultSheetDataModel modeladd = new FarPoint.Web.Spread.Model.DefaultSheetDataModel(dt);

            sv.DataModel = modeladd;
            sv.FpSpread.DataBind();


            int col = sv.ColumnHeader.Columns.Count - 1;
            sv.ColumnHeader.Cells[row, col].Text = sHeaderNM;
            sv.ColumnHeader.Columns[col].DataField = sDataFild;
            sv.Columns[col].DataField = dt.Columns[col].ColumnName;
            sv.Columns[col].Locked = Lock;
            sv.ColumnHeader.Columns[col].Width = Width;

            sv.Columns[col].HorizontalAlign = hAlign;

            sv.Columns[col].CellType = CellType;

            if(!Visible)
            {
                sv.Columns[col].Visible = Visible;
            }
        }

        public static Cell GetCell(SheetView sv, int row, string DataFild)
        {
            int col=0;
            bool chk = false;

            for (int i = 0; i < sv.ColumnCount; i++ )
            {
                if(sv.Columns[i].DataField == DataFild)
                {
                    col = i;

                    chk = true;
                    break;
                }
            }

            if (chk)
            {
                return sv.Cells[row, col];
            }
            else
            {
                return null;
            }
        }

        public static String GetColName(SheetView sv, string DataFild )
        {
            string sResult = "";

            int row = sv.ColumnHeader.RowCount - 1;

            for (int i = 0; i < sv.ColumnCount; i++ )
            {
                if(sv.Columns[i].DataField == DataFild)
                {
                    sResult = sv.ColumnHeader.Cells[row, i].Text;
                }
            }

            return sResult;
        }

        public static bool CheckValidationSV(SheetView sv, string ChkFild, Page MyPage)
        {
             bool result = true;
            string[] arrChkF = ChkFild.Split(',');

            DataSet ds = (DataSet)sv.DataSource;

            ds.Tables[0].DefaultView.RowFilter = "CHK='1' ";
            DataTable dt = ds.Tables[0].DefaultView.ToTable();

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < arrChkF.Length; i++)
                {

                    string sSelect = "";

                    if (dt.Columns[arrChkF[i].Trim()].DataType == typeof(DateTime))
                    {
                        sSelect = arrChkF[i].Trim() + " IS NULL";
                    }
                    else
                    {
                        sSelect = arrChkF[i].Trim() + " IS NULL OR " + arrChkF[i].Trim() + " = '' ";
                    }

                    DataRow[] dr = dt.Select(sSelect);

                    if (dr.Length > 0)
                    {
                        MessageBox.ShowMessage(GetColName(sv, arrChkF[i].Trim()) + "는 필수 입력 값 입니다.", MyPage);
                        result = false;
                        break;
                    }
                }
            }
             return result;
        }

        public static DataTable GetCheckDataTable(SheetView sv)
        {
            DataSet ds = (DataSet)sv.DataSource;

            ds.Tables[0].DefaultView.RowFilter = "CHK='1' ";
            DataTable dt = ds.Tables[0].DefaultView.ToTable();

            return dt;
        }

        public static void ExportExcel(FpSpread fp, Page MyPage)
        {
            int index = -1;

            if (fp.Sheets[0].Rows.Count > 0)
            {

                //for (int i = 0; i < fp.Sheets.Count; i++)
                //{
                //if (!fp.Sheets[i].Visible)
                //{
                //    fp.Sheets[i].Visible = true;
                //    index = i;

                string dt = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + "_" + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00");

                System.IO.MemoryStream m_stream = new System.IO.MemoryStream();


                fp.SaveExcel(m_stream, FarPoint.Excel.ExcelSaveFlags.SaveBothCustomRowAndColumnHeaders);
                m_stream.Position = 0;

                MyPage.Response.ContentType = "application/vnd.ms-excel";
                MyPage.Response.AddHeader("content-disposition", "inline; filename=" + dt + ".xls");
                MyPage.Response.BinaryWrite(m_stream.ToArray());
                //Response.BinaryWrite(m_stream2.ToArray());
                MyPage.Response.End();

                


                //System.IO.MemoryStream s = new System.IO.MemoryStream();
                //fp.SaveExcel(s);
                //s.Position = 0;
                //fp.OpenExcel(s);
                //s.Close();

                //    if (index > -1)
                //    {
                //        fp.Sheets[index].Visible = false;
                //    }

                //}
                //}

            }
        }

	}