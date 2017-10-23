using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace T_IF_RCV_PROD_ORD_KO441
{
    public partial class FormMain : Form
    {
        /* 프로그램 목적 : MES <-> ERP간 작업지시, 라우팅, 자재입출고 정보를 I/F 하기 위해 각각의 DB에 Connect하여 정보를 갱신한다.
         * 작성일자 : 2015-01-16
         * 작성자 : yoosr
         * 수정내역 : 
         * */

        // Global variable
        T_IF_RCV_PROD_ORD_KO441.DBConn oraConn = null;      // MES DB Connection (Oracle)
        T_IF_RCV_PROD_ORD_KO441.DBConn mssqlConn = null;    // ERP DB Connection (MS-Sql)

        private string mesConnectionString = "Provider=MSDAORA;Data Source=CCUBE;user id=mighty;password=mighty";
        private string erpConnectionString = "Provider=SQLOLEDB;Data Source=192.168.10.15;Initial Catalog=NEPES;uid=sa;pwd=nepes01!";

        #region From Initial
        public FormMain()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // #01.Initial
            btnTimer.BackColor = Color.Red;
            btnTimer.Text = "Timer Start";

            // #02.Initial Timer
            mainTimer.Enabled = false;
            //mainTimer.Interval = 50000; //별도의 아이콘에 Interval을 지정되어 있음

        }
        #endregion

        #region Timer Event
        /// <summary>
        /// Timer Tick
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mainTimer_Tick(object sender, EventArgs e)
        {
            mainTimer.Enabled = false;

            Send_Interface_Data();

            /*
            // 20150115, ERP에서 입고 잡은 PO기준으로 MCS 입고 테이블에 반영 해준다.
            Send_MCS_Data();
             * 주석날짜 : 2015.12.02, 최초 작성자 : 강중구D, 수정자 : 안창주
             * 상기 함수 주석사용 : 
                실제 I/F EXE파일에서 사용되는 메일링 시스템이 되지 않기에 확인을 위해 프로그램 확인 함, 프로그램 확인 중 MCS 함수내에서 사용되는 쿼리 및 변수대입 시 기존부터 오류가 있어 Catch문으로 떨어짐을 확인.
                허나, 메일링 시스템이 되지 않기에 별도의 메일 없이 타이머기능으로 작업지시만 지속적으로 반영되고 있었음. (작업지시 먼저 반영 후 MCS가 반영되게끔 프로그램이 코딩되어 있음)
                최초 작성자 및 MCS담당자 확인 시 정확한 이력을 파악할 수 없었음.
                15년 1월 이후에 MCS테이블 업데이트가 없었으나 별다른 확인요청이 없어 주석처리를 하기로 결정.
                (실제 MCS에서 사용되는 테이블의 Flag를 보면 15년1월 이후에는 전부 'N')
            */

            mainTimer.Enabled = true;
        }

        private void btnTimer_Click(object sender, EventArgs e)
        {
            if (btnTimer.BackColor == Color.Red)
            {
                btnTimer.BackColor = Color.Green;
                btnTimer.Text = "Timer End";

                //mainTimer.Interval = 1800000; //별도의 아이콘에 Interval을 지정되어 있음
                mainTimer.Interval = 1000 * 180;

                mainTimer.Enabled = true;

                mainTimer_Tick(null, null);
            }
            else
            {
                btnTimer.BackColor = Color.Red;
                btnTimer.Text = "Timer Start";

                mainTimer.Enabled = false;
            }
        }
        #endregion

        #region DB 접속 테스트
        /// <summary>
        /// btnMSSql_Click - MS-SQL 접속 테스트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMSSql_Click(object sender, EventArgs e)
        {
            try
            {
                if (mssqlConn == null)
                    mssqlConn = new DBConn(erpConnectionString);

                this.txtMSSql.Text = "Connect OK!!";
            }
            catch (Exception ex)
            {
                //gridHistory.Rows.Insert(0, new object[] { "SHIP", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "ERROR", ex.Message });
                this.txtMSSql.Text = "Connect Fail!! [" + ex.ToString() + "]";
                return;
            }
        }

        /// <summary>
        /// btnOracle_Click - Oracle 접속 테스트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOracle_Click(object sender, EventArgs e)
        {
            try
            {
                if (oraConn == null)
                    oraConn = new DBConn(mesConnectionString);

                this.txtOracle.Text = "Connect OK!!";
            }
            catch (Exception ex)
            {
                //gridHistory.Rows.Insert(0, new object[] { "SHIP", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "ERROR", ex.Message });
                this.txtOracle.Text = "Connect Fail!! [" + ex.ToString() + "]";
                return;
            }
        }
        #endregion

        #region MES -> ERP, 작업지시 연동
        /// <summary>
        /// Send_Interface_Data - 작업지시 연동 메인
        /// </summary>
        private void Send_Interface_Data()
        {
            try
            {
                if (mssqlConn == null)
                    mssqlConn = new DBConn(erpConnectionString);

                if (oraConn == null)
                    oraConn = new DBConn(mesConnectionString);

                DataTable dtMesData = GetMESData();

                string erpOrderNo = string.Empty;
                string mesOrderNo = string.Empty;

                if (dtMesData != null && dtMesData.Rows.Count > 0)
                {
                        //string sToUser = "50458@nepes.co.kr";
                        string sToUser = "ahncj@nepes.co.kr";

                        string sSubject = "[NEPES][" + DateTime.Now.ToString("yyyyMMdd") + "] MES->ERP 작업지시 I/F ERROR REPORT";
                        string strMailData = "";

                        strMailData += "  This mail is sent by nepes auto sending mail program.";
                        strMailData += "  <BR>";
                        strMailData += "  <BR>MES작업지시 생성 테이블(T_IF_RCV_PROD_ORD_KO441)에 ITEM_CD가 NULL인 값이 존재합니다.<BR>";
                        strMailData += "  <BR><BR>";    
                        strMailData += "  <table border='1'>";
                        strMailData += "  <tr><th> PRODT_ORDER_NO </th><th> PLANT_CD </th> ";

                        for (int rowIdx = 0; rowIdx < dtMesData.Rows.Count; rowIdx++)
                        {
                            // ITEM_CD가 NULL인건 알람메일 전송하고 연동대상에서 제외한다.
                            if (string.IsNullOrEmpty(dtMesData.Rows[rowIdx]["ITEM_CD"].ToString()))
                            {                                
                                strMailData += " <tr><td>" + dtMesData.Rows[rowIdx]["PRODT_ORDER_NO"].ToString() + "</td><td>" + dtMesData.Rows[rowIdx]["PLANT_CD"].ToString() + "</td> ";
                                continue;
                            }

                            //OrderNO가 있을 경우 연동대상에 포함한다
                            if (string.IsNullOrEmpty(erpOrderNo))
                            {
                                erpOrderNo = "''" + dtMesData.Rows[rowIdx]["PRODT_ORDER_NO"].ToString() + "''";
                                mesOrderNo = "'" + dtMesData.Rows[rowIdx]["PRODT_ORDER_NO"].ToString() + "'";
                            }
                            else
                            {
                                erpOrderNo += ",''" + dtMesData.Rows[rowIdx]["PRODT_ORDER_NO"].ToString() + "''";
                                mesOrderNo += ",'" + dtMesData.Rows[rowIdx]["PRODT_ORDER_NO"].ToString() + "'";
                            }
                        }

                        //50분 단위로 돌기에 07시에는 메일링이 갈 수 있게 한다
                        //if (DateTime.Now.ToString("HH") == "07")
                        //{
                            strMailData += "  </table> ";
                            CommonFunction.SendMail("nepes_sys@nepes.co.kr", "시스템관리자", sToUser, "", sSubject, strMailData, "", "", true, false);
                        //}

                        //모두 다 연동대상에서 제외될 경우 erpOrderNo 변수에는 값이 없으므로 체크 후 함수를 호출한다
                        if (!string.IsNullOrEmpty(erpOrderNo))
                        {
                            SetERPData(erpOrderNo);
                            SetMESData(mesOrderNo);
                        }
                }
            }
            catch (Exception ex)
            {
                // Send Mail
                string sSubject = "[NEPES][" + DateTime.Now.ToString("yyyyMMdd") + "] MES->ERP 작업지시 I/F ERROR REPORT";
                string strMailData = "";
                strMailData += "  This mail is sent by nepes auto sending mail program.";
                strMailData += "  <BR>";
                strMailData += "  <BR>MES -> ERP 작업지시 연동에 오류가 발생했습니다. <BR>";
                strMailData += "  <BR>세부메시지 : " + ex.Message + "<BR>";
                strMailData += "  <BR>";
                strMailData += "  <BR>MES의 'T_IF_RCV_PROD_ORD_KO441' 테이블 확인 바랍니다.<BR>";
                strMailData += "  <BR>";                

                string sToUser = "50458@nepes.co.kr";
                //string sToUser = "ahncj@nepes.co.kr";

                CommonFunction.SendMail("nepes_sys@nepes.co.kr", "시스템관리자", sToUser, "", sSubject, strMailData, "", "", true, false);
            }
            finally
            {
                mssqlConn.DisconnectDBConn();
                oraConn.DisconnectDBConn();
            }
        }

        /// <summary>
        /// GetMESData - MES->ERP로 전달할 작업지시 리스트를 가져온다.
        /// </summary>
        /// <returns></returns>
        private DataTable GetMESData()
        {
            DataTable dtResult = null;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("SELECT PRODT_ORDER_NO, NVL(ITEM_CD, '') AS ITEM_CD, PLANT_CD ");
            //MODIFY BY SLYOO : 2015-01-28 : MCS 연동으로 인한 Connection 변경
            sb.AppendLine("  FROM UNIERPSEMI.T_IF_RCV_PROD_ORD_KO441 ");
            sb.AppendLine(" WHERE ERP_APPLY_FLAG1 = 'N' ");
            sb.AppendLine(" ORDER BY PRODT_ORDER_NO ");

            dtResult = oraConn.ExecuteQuery(sb.ToString());

            return dtResult;
        }

        /// <summary>
        /// SetERPData - ERP로 전송
        /// </summary>
        /// <param name="sProdtOrderNoList"></param>
        /// <returns></returns>
        private int SetERPData(string sProdtOrderNoList)
        {
            /*
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("DECLARE @STRQUERY   VARCHAR(2084) ");
            sb.AppendLine("SELECT @STRQUERY = NULL ");
            sb.AppendLine("SELECT  @STRQUERY = ' ");
            sb.AppendLine("INSERT INTO T_IF_RCV_PROD_ORD_KO441 ");
            sb.AppendLine("SELECT A.* FROM OPENQUERY(CCUBE,''' + ' ");
            sb.AppendLine("     SELECT * FROM UNIERPSEMI.T_IF_RCV_PROD_ORD_KO441 WHERE ERP_APPLY_FLAG1 = ''''N'''' ");
            sb.AppendLine("        AND PRODT_ORDER_NO IN (" + sProdtOrderNoList + ") ");
            sb.AppendLine(" '') A ' ");
            sb.AppendLine("EXEC (@STRQUERY) ");
            //usp_interface_mes_to_erp
            int nResult = oraConn.ExecuteNonQuery(sb.ToString());
             * */
            string sQuery = "INSERT INTO T_IF_RCV_PROD_ORD_KO441 SELECT A.* FROM OPENQUERY(CCUBE,' SELECT * FROM UNIERPSEMI.T_IF_RCV_PROD_ORD_KO441 WHERE ERP_APPLY_FLAG1 = ''N'' AND PRODT_ORDER_NO IN (" + sProdtOrderNoList + ") ') A";
            
            int nResult = mssqlConn.ExecuteNonQuery(sQuery);

            return nResult;
        }

        /// <summary>
        /// SetMESData - 전송 완료된 작업지시에 대해 MES에 상태 Update.
        /// </summary>
        /// <param name="sProdtOrderNoList"></param>
        /// <returns></returns>
        private int SetMESData(string sProdtOrderNoList)
        {
            StringBuilder sb = new StringBuilder();
            //MODIFY BY SLYOO : 2015-01-28 : MCS 연동으로 인한 Connection 변경
            sb.AppendLine("UPDATE UNIERPSEMI.T_IF_RCV_PROD_ORD_KO441 ");
            //sb.AppendLine("UPDATE T_IF_RCV_PROD_ORD_KO441 ");
            sb.AppendLine("SET ERP_APPLY_FLAG1 = 'Y' ");
            sb.AppendLine("WHERE ERP_APPLY_FLAG1 = 'N' AND PRODT_ORDER_NO IN (" + sProdtOrderNoList + ") ");

            int nResult = oraConn.ExecuteNonQuery(sb.ToString());

            return nResult;
        }

        private void btnManual_Click(object sender, EventArgs e)
        {
            Send_Interface_Data();
        }
        #endregion


        /*
         *
         
             * 주석날짜 : 2015.12.02, 최초 작성자 : 강중구D, 수정자 : 안창주
             * 상기 함수 주석사용 : 
                실제 I/F EXE파일에서 사용되는 메일링 시스템이 되지 않기에 확인을 위해 프로그램 확인 함, 프로그램 확인 중 MCS 함수내에서 사용되는 쿼리 및 변수대입 시 기존부터 오류가 있어 Catch문으로 떨어짐을 확인.
                허나, 메일링 시스템이 되지 않기에 별도의 메일 없이 타이머기능으로 작업지시만 지속적으로 반영되고 있었음. (작업지시 먼저 반영 후 MCS가 반영되게끔 프로그램이 코딩되어 있음)
                최초 작성자 및 MCS담당자 확인 시 정확한 이력을 파악할 수 없었음.
                15년 1월 이후에 MCS테이블 업데이트가 없었으나 별다른 확인요청이 없어 주석처리를 하기로 결정.
                (실제 MCS에서 사용되는 테이블의 Flag를 보면 15년1월 이후에는 전부 'N') 
         
         
        #region ERP -> MCS, 생산입고 연동
        /// <summary>
        /// Send_MCS_Data - 생산입고 메인
        /// </summary>
        private void Send_MCS_Data()
        {
            try
            {
                if (mssqlConn == null)
                    mssqlConn = new DBConn(erpConnectionString);

                if (oraConn == null)
                    oraConn = new DBConn(mesConnectionString);

                DataTable dtPOList = GetPOList();

                string strPoNo = string.Empty;
                string strPoSeqNo = string.Empty;
                string strItemNo = string.Empty;
                int nQty = 0;
                string strGrDate = string.Empty;

                if (dtPOList != null && dtPOList.Rows.Count > 0)
                {
                    for (int rowIdx = 0; rowIdx < dtPOList.Rows.Count; rowIdx++)
                    {
                        strPoNo = dtPOList.Rows[rowIdx]["PO_NO"].ToString();
                        strPoSeqNo = dtPOList.Rows[rowIdx]["PO_SEQ_NO"].ToString();
                        strItemNo = dtPOList.Rows[rowIdx]["ITEM_CD"].ToString();
                        /nQty = Convert.ToInt32(dtPOList.Rows[rowIdx]["GR_QTY"].ToString());
                        strGrDate = dtPOList.Rows[rowIdx]["GR_DT"].ToString();

                        SetMCSData(strPoNo, strPoSeqNo, strItemNo, nQty, strGrDate);
                    }
                }
            }
            catch (Exception ex)
            {
                // Send Mail
                string sSubject = "[NEPES][" + DateTime.Now.ToString("yyyyMMdd") + "] ERP->MCS 생산입고 I/F ERROR REPORT";
                string strMailData = "";
                strMailData += "  This mail is sent by nepes auto sending mail program.";
                strMailData += "  <BR>";
                strMailData += "  <BR>ERP -> MCS 생산입고 연동에 오류가 발생했습니다. <BR>";
                strMailData += "  <BR>세부메시지 : " + ex.Message + "<BR>";
                strMailData += "  <BR>";

                string sToUser = "50458@nepes.co.kr";
                //string sToUser = "ahncj@nepes.co.kr";
                CommonFunction.SendMail("nepes_sys@nepes.co.kr", "시스템관리자", sToUser, "", sSubject, strMailData, "", "", true, false);
            }
            finally
            {
                mssqlConn.DisconnectDBConn();
                oraConn.DisconnectDBConn();
            }
        }

        /// <summary>
        /// GetPOList - ERP에서 생산입고시 생성된 I/F PO정보 조회
        /// </summary>
        /// <returns></returns>
        private DataTable GetPOList()
        {
            DataTable dtResult = null;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("SELECT PO_NO, PO_SEQ_NO, CONVERT(VARCHAR(18), GR_DT, 112) AS GR_DT, ITEM_CD, convert(int, GR_QTY) AS GR_QTY ");
            sb.AppendLine("  FROM T_IF_SND_PO_TO_MCS_KO441 ");
            sb.AppendLine(" WHERE MES_APPLY_FLAG = 'N' ");
            sb.AppendLine(" ORDER BY PO_NO, PO_SEQ_NO, ITEM_CD, GR_DT ");

            dtResult = mssqlConn.ExecuteQuery(sb.ToString());

            return dtResult;
        }

        /// <summary>
        /// SetMESData, ERP에서 입고잡은 PO를 수량기준으로 MCS에서 입고처리 해준다.
        /// </summary>
        /// <param name="sProdtOrderNoList"></param>
        /// <returns></returns>
        private int SetMCSData(string strPoNo, string strPoSeqNo, string strItemNo, int nQty, string strGrDT)
        {
            DataTable dtResult = null;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("SELECT MAT_LOT_NO, IMPORT_QTY ");
            sb.AppendLine("  FROM AMAT_MAT_STORE_HISTORY ");
            sb.AppendLine("WHERE PO_NO = '" + strPoNo + "' ");
            sb.AppendLine("  AND ITEM_CODE = '" + strItemNo + "' ");
            sb.AppendLine("  AND ERP_UPDATE_TIME = '99991231235959' ");
            sb.AppendLine("  AND TRANS_CODE = 'MSIN' ");
            sb.AppendLine("ORDER BY PO_NO, IMPORT_DATE, MAT_LOT_NO ");

            dtResult = oraConn.ExecuteQuery(sb.ToString());

            string strMatLotNo = string.Empty;

            for (int nRowIdx = 0; nRowIdx < dtResult.Rows.Count; nRowIdx++)
            {
                int mesImportQty = Convert.ToInt32(dtResult.Rows[nRowIdx]["IMPORT_QTY"]);
                //int mesImportQty = Convert.ToInt32(dtResult.Rows[nRowIdx]["IMPORT_QTY"].ToString());

                if (nQty >= mesImportQty)
                {
                    nQty = nQty - mesImportQty;

                    if (string.IsNullOrEmpty(strMatLotNo))
                    {
                        strMatLotNo = "'" + dtResult.Rows[nRowIdx]["MAT_LOT_NO"].ToString() + "'";
                    }
                    else
                    {
                        strMatLotNo += ",'" + dtResult.Rows[nRowIdx][0].ToString() + "'";
                    }
                }
                else
                {
                    break;
                }
            }

            int nResult = 0;

            if (!string.IsNullOrEmpty(strMatLotNo))
            {
                sb = new StringBuilder();
                sb.AppendLine("UPDATE AMAT_MAT_STORE_HISTORY ");
                sb.AppendLine("SET ERP_UPDATE_TIME = '" + strGrDT + "000000'");
                sb.AppendLine("WHERE PO_NO = '" + strPoNo + "' ");
                sb.AppendLine("  AND ITEM_CODE = '" + strItemNo + "' ");
                sb.AppendLine("  AND MAT_LOT_NO IN (" + strMatLotNo + ") ");
                sb.AppendLine("  AND TRANS_CODE = 'MSIN' ");

                nResult = oraConn.ExecuteNonQuery(sb.ToString());

                SetPOData(strPoNo, strPoSeqNo, strItemNo, nQty, strGrDT);
            }

            return nResult;
        }

        /// <summary>
        /// ERP I/F 테이블에 처리결과 Update
        /// </summary>
        /// <param name="strPoNo"></param>
        /// <param name="strItemNo"></param>
        /// <param name="nQty"></param>
        /// <param name="strGrDT"></param>
        /// <returns></returns>
        //private int SetPOData(string strPoNo, string strPoSeqNo, string strItemNo, int nQty, string strGrDT)
        private int SetPOData(string strPoNo, string strPoSeqNo, string strItemNo, int nQty, string strGrDT)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("UPDATE T_IF_SND_PO_TO_MCS_KO441 ");
            sb.AppendLine("SET MES_APPLY_FLAG = 'Y', ");
            sb.AppendLine("    MES_RECEIVE_FLAG = 'Y', ");
            sb.AppendLine("    MES_RECEIVE_DT = GETDATE() ");
            //sb.AppendLine("WHERE PO_NO = '" + strPoNo + "' AND PO_SEQ_NO = '" + strPoSeqNo + "' AND ITEM_CD = '" + strItemNo + "' AND GR_QTY = '" + nQty.ToString() + "' AND CONVERT(VARCHAR(8), GR_DT, 112) = '" + strGrDT + "' ");
            sb.AppendLine("WHERE PO_NO = '" + strPoNo + "' AND PO_SEQ_NO = '" + strPoSeqNo + "' AND ITEM_CD = '" + strItemNo + "' AND GR_QTY = '" + nQty + "' AND CONVERT(VARCHAR(8), GR_DT, 112) = '" + strGrDT + "' ");

            int nResult = mssqlConn.ExecuteNonQuery(sb.ToString());

            return nResult;
        }
        #endregion

        */
    }
}
