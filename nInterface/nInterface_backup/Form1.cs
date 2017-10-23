using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections;

namespace nInterface
{
    public partial class Form1 : Form
    {

        //처리해야 할 예외 case
        //1. IF시 이미 들어간 데이터가 있을 경우
        //2. 테이블 SELECT 및 INSERT시 Table Lock에 의해 타임아웃이 발생될 경우


        #region Member variable declarations(location: App.config)
        /// <summary>
        /// 인터페이스 할 쿼리정보입니다.
        /// </summary>
        private static readonly string _initPath = System.Configuration.ConfigurationSettings.AppSettings["INIT_PATH"];

        /// <summary>
        /// 로그를 저장할 위치입니다.
        /// </summary>
        private static readonly string _LogPath = System.Configuration.ConfigurationSettings.AppSettings["LOG_PATH"];

        /// <summary>
        /// 데이터베이스 접속을 보관하는 파일위치 입니다.
        /// </summary>
        private static readonly string _dbPart = System.Configuration.ConfigurationSettings.AppSettings["DB_PATH"];

        /// <summary>
        /// 오류발생 시 발송 될 메일주소와 휴대폰주소 정보위치 입니다.
        /// </summary>
        private static readonly string _setPath = System.Configuration.ConfigurationSettings.AppSettings["SET_PATH"];

        /// <summary>
        /// ini파일에서 섹션을 기준으로 최종 데이터를 담은 리스트string 객체 입니다. (ERP, MES, SET정보)
        /// </summary>
        List<List<string>> _iniERPInfo = new List<List<string>>();
        List<List<string>> _iniMESInfo = new List<List<string>>();
        List<List<string>> _iniSetInfo = new List<List<string>>();

        /// <summary>
        /// 발생된 애러를 별도로 참아 사용자에게 전송하는 리스트string 객체 입니다.
        /// </summary>
        List<string> _errorThrow = new List<string>();
        #endregion

        #region It Generate an Database object that is responsible for connection (ERP, MES)
        /// <summary>
        /// DB객체입니다 (MES, ERP)
        /// </summary>
        static nInterface.DBConn _DBConn = new DBConn();
        #endregion

        #region Other variables.
        static iniFileHandle.Win32Reg Win32Reg = null;
        static nInterface.LogControl logCtrl = new LogControl(_LogPath);
        static nInterface.CommonFunction comFunc = new CommonFunction();
        #endregion

        /// <summary>
        /// it is the program initial entry point
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// And proceeds to initialize before programming begins.
        /// (db connection information or  TextBox String initialization)
        /// </summary>
        private void initSetting()
        {
            //프로각종 경로 라벨 셋팅 
            lblPath.Text = _dbPart;
            lbldbQuery.Text = _initPath;
            lbldbLog.Text = _LogPath;
            label9.Text = _setPath;

            mainTimer.Enabled = false;
        }

        /// <summary>
        /// ERP DB 접속을 테스트합니다.
        /// </summary>
        private void ERPConnection_Click(object sender, EventArgs e)
        {
            try
            {
                if (_DBConn.ConnectionERPString == null || _DBConn.ConnectionERPString.Equals(""))
                {
                    _DBConn.ConnectionERPString = System.IO.File.ReadAllText(_dbPart + "ERPInfo.ini").ToString();
                    _DBConn.InitERP_DBConn();
                }

                txtERPDB.Text = "Connection";
            }
            catch (Exception ex)
            {
                setLogHandler("Connection..", ex, "", "", true, true);
            }
        }

        /// <summary>
        /// MES DB 접속을 테스트합니다.
        /// </summary>
        private void MESConnection_Click(object sender, EventArgs e)
        {
            try
            {
                if (_DBConn.ConnectionMESString == null || _DBConn.ConnectionMESString.Equals(""))
                {
                    _DBConn.ConnectionMESString = System.IO.File.ReadAllText(_dbPart + "MESInfo.ini");
                    _DBConn.InitMES_DBConn();
                }
                
                txtMESDB.Text = "Connection";
            }
            catch (Exception ex)
            {
                setLogHandler("Connection..", ex, "", "", true, true);
            }
        }

        /// <summary>
        /// 쿼리가 저장된 ini파일을 파싱합니다.
        /// </summary>
        private void SetFileStream()
        {
            string[] iniSection = null;

            if (_iniERPInfo.Count == 0)
            {
                Win32Reg = new iniFileHandle.Win32Reg(_initPath + "ERP-MES.ini");
                iniSection = Win32Reg.GetSectionNames();
                SetiniInfo(iniSection, "ERP");
                SetlistControl(iniSection, "ERP");     //listView 셋팅 (배열이 초기화 되기 전 셋팅)
            }

            if (_iniMESInfo.Count == 0)
            {
                Win32Reg.szFileName = _initPath + "MES-ERP.ini";
                //Win32Reg = new iniFileHandle.Win32Reg(_initPath + "MES-ERP.ini");
                iniSection = Win32Reg.GetSectionNames();
                SetiniInfo(iniSection, "MES");
                SetlistControl(iniSection, "MES");     //listView 셋팅 (배열이 초기화 되기 전 셋팅)
            }

            if (_iniSetInfo.Count == 0)
            {
                Win32Reg = new iniFileHandle.Win32Reg(_setPath + "setInfo.ini");
                iniSection = Win32Reg.GetSectionNames();
                SetiniInfo(iniSection, "SET");
            }

        }

        /// <summary>
        /// 전역변수(list)에 2차원 배열로 데이터를 대입한다 (프로그램을 종료하지 않는 한 계속 사용되는 변수)
        /// </summary>
        /// <param name="iniInfo">전역변수로 대입시킬 ini파일의 정보가 대입 된 배열입니다.</param>
        /// <param name="Gubun">처리를 확인하는 구분자 입니다.</param>
        private void SetiniInfo(string[] iniInfo, string Gubun)
        {
            string[] tempArray;

            for (int index = 0; index < iniInfo.Length; index++)
            {
                if (Gubun.ToUpper() == "ERP")
                    _iniERPInfo.Add(new List<string>());
                else if (Gubun.ToUpper() == "MES")
                    _iniMESInfo.Add(new List<string>());
                else
                    _iniSetInfo.Add(new List<string>());


                tempArray = Win32Reg.GetPairsBySection(iniInfo[index]);

                for (int index2 = 0; index2 < tempArray.Length; index2++)
                {
                    if (Gubun.ToUpper() == "ERP")
                        _iniERPInfo[index].Add(tempArray[index2]);
                    else if (Gubun.ToUpper() == "MES")
                        _iniMESInfo[index].Add(tempArray[index2]);
                    else
                        _iniSetInfo[index].Add(tempArray[index2]);
                }
            }

        }

        /// <summary>
        /// If you click the Start button it is the initial entry point.
        /// </summary>
        private void Start_Click(object sender, EventArgs e)
        {
            try
            {
                //프로그램 기준정보 셋팅 및 최초 접속 Connection 확인
                bool state = SetBaseInfo();

                //기준정보가 false일 경우 return으로 프로그램을 실행하지 않는다.
                if (state == false)
                    return;

                //타이머시간 셋팅(밀리초이기 때문에 분을 만들기 위해서 화면에서 전달받은 숫자 * 60초를 해야 함 (1초 = 1000초))
                mainTimer.Interval = Convert.ToInt32(cbTimer.Text) * 60000;
                mainTimer_Tick(null, null);
            }
            catch (Exception ex)
            {
                setLogHandler("Start_Click..", ex, "", "", true, true);
            }
        }

        /// <summary>
        /// begin Program Process.
        /// </summary>
        private void processStart()
        {
            //타이머를 실행할 때 마다 최종시간을 화면에 표시
            lblTime.Text = System. DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss");
            
            //mes interface 로직 실행
            mesDataProcess();

            //erp interface 로직 실행
            erpDataProcess();
        }

        /// <summary>
        /// If you click the Stop button it is the initial entry point.
        /// </summary>
        private void Stop_Click(object sender, EventArgs e)
        {
            setControlLock("E");

            if (MessageBox.Show("프로그램을 종료하시겠습니까?", "종료확인", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                //DB연결 종료
                _DBConn.CloseERPDBConn();
                _DBConn.CloseMESDBConn();

                //프로그램 종료
                Environment.Exit(0);
            }
        }

        /// <summary>
        /// 배열 index로 프로그래밍을 할 경우 ini 파일에서 순서가 바뀔 경우 문제가 발생될 수 있어 해쉬테이블로 전환
        /// </summary>
        /// <param name="hst">해쉬테이블 객체입니다.</param>
        /// <param name="Key">해쉬테이블의 Key값 입니다.</param>
        /// <param name="index">2차원 배열 중 첫번째 index 값 입니다.</param>
        /// <param name="index2">2차원 배열 중 두번째 index 값 입니다.</param>
        /// <param name="Gubun">메소드사용을 구분하는 구분자입니다. (ERP, MES, SET)</param>
        private void SethashTable(Hashtable hst, string Key, int index, int index2, string Gubun)
        {
            string tempInfo = string.Empty;

            if (Gubun.ToUpper() == "ERP")
                tempInfo = _iniERPInfo[index][index2].ToString();
            else if (Gubun.ToUpper() == "MES")
                tempInfo = _iniMESInfo[index][index2].ToString();
            else
                tempInfo = _iniSetInfo[index][index2].ToString();

            //앞에 구분자로 필요없는 부분을 자름 ("=" 이후부터 사용되기 때문에 +1 을 실시) 
            hst.Add(Key, tempInfo.Substring(tempInfo.IndexOf("=") + 1));
        }

        /// <summary>
        /// MES interface 처리로직 입니다. (MES -> ERP)
        /// </summary>
        private void mesDataProcess()
        {
            DataTable dtResult = new DataTable();
            Hashtable hstResut = new Hashtable();
            string tempIniKey = string.Empty;
            string Gubun = "MES";

            for (int index = 0; index < _iniMESInfo.Count(); index++)
            {
                for (int index2 = 0; index2 < _iniMESInfo[index].Count(); index2++)
                {
                    tempIniKey = _iniMESInfo[index][index2].Substring(0, _iniMESInfo[index][index2].IndexOf("="));
                    SethashTable(hstResut, tempIniKey, index, index2, Gubun);
                }

                //메인 처리 메소드 (DB select 및 트랜잭션 발생)
                mainProcessing(hstResut, Gubun);

                //다음 로직을 진행하기 전 해쉬테이블을 초기화
                hstResut.Clear();
            }

        }

        /// <summary>
        /// ERP interface 처리로직 입니다. (ERP -> MES)
        /// </summary>
        private void erpDataProcess()
        {
            DataTable dtResult = new DataTable();
            Hashtable hstResut = new Hashtable();
            string tempIniKey = string.Empty;
            string Gubun = "ERP";

            for (int index = 0; index < _iniERPInfo.Count(); index++)
            {
                for (int index2 = 0; index2 < _iniERPInfo[index].Count(); index2++)
                {
                    tempIniKey = _iniERPInfo[index][index2].Substring(0, _iniERPInfo[index][index2].IndexOf("="));
                    SethashTable(hstResut, tempIniKey, index, index2, Gubun);
                }

                //메인 처리 메소드 (DB select 및 트랜잭션 발생)
                mainProcessing(hstResut, Gubun);
                
                //다음 로직을 진행하기 전 해쉬테이블을 초기화
                hstResut.Clear();
            }

        }

        /// <summary>
        /// form의 시작점 입니다.
        /// </summary>
        private void Form1_Load(object sender, EventArgs e)
        {
            initSetting();
        }

        /// <summary>
        /// 데이터를 처리하는 로직입니다. (DB 트랜잭션 발생)
        /// </summary>
        /// <param name="hst">처리해야 할 데이터(쿼리)가 들어있는 해쉬테이블 정보입니다.</param>
        /// <param name="Gubun">메소드사용을 구분하는 구분자입니다. (ERP, MES, SET)</param>
        private void mainProcessing(Hashtable hst, string Gubun)
        {
            string querySource = hst["SOURCE"].ToString();      //SELECT 쿼리 (대상 조회)
            string queryTarget = string.Empty;
            string queryTransaction = string.Empty;
            string queryTransactionUpdate = string.Empty;
            string queryUpdate = string.Empty;
            string targetOption = hst["TARGET_OPTION"].ToString().ToUpper();         //insert 구분인지 update 구문인지를 구별 할 수 있는 ini 구분자
            int cnt = 0;
            bool transactionCheck;

            DataTable dtResult = new DataTable();

            try
            {
                //INSERT 할 대상 조회 (SELECT)
                setLogHandler(Gubun, null, querySource, "", false, false);
                dtResult = _DBConn.ExecuteQuery(querySource, Gubun.ToUpper());

                cnt = dtResult.Rows.Count;

                //데이터테이블을 로우로 변환 해 각 Row별 처리로직
                //문제가 발생되는 Row는 건너뛰고 다음Row를 실행해야 하기 때문에 일괄처리를 할 수 없음 (1~10번까지 중 8번째는 sql오류가 발생한다고 하면 1~7번, 9~10번은 처리를 정상적으로 해야하기 때문에..)
                while (cnt > 0)
                {
                    queryTransaction = "";
                    transactionCheck = false;   
                    //실제 트랜젝션이 일어날 때 오류가 발생되었어도 해당 쿼리는 건너뛰고 다음쿼리를 진행하기 위해 중첩 try-catch 사용 
                    //첫 번째 try-catch의 용도 : SELECT문이 오류가 있어도 다음 쿼리를 실행하기 위한 예외처리
                    //두 번째 try-catch의 용도 : SELECT된 데이터를 트랜젝션 쿼리가 발생될 때 문제부분을 건너뛰고 실행하기 위함
                    try
                    {
                        queryTransaction = getQuery(hst, dtResult.Rows[cnt - 1], Gubun.ToUpper(), targetOption);
                        queryTransactionUpdate = getUpdtQuery(hst, dtResult.Rows[cnt - 1], Gubun.ToUpper());
                        
                        //조합된 insert, update 쿼리를 실행하기 전 로그 기록
                        setLogHandler(Gubun, null, "Transaction_Start", queryTransaction, false, false);
                        setLogHandler(Gubun, null, "Transaction_Start", queryTransactionUpdate, false, false);

                        transactionCheck = _DBConn.ExecuteTransactionNonQuery(queryTransaction, queryTransactionUpdate, Gubun.ToUpper());

                        if (transactionCheck == true)
                            setLogHandler(Gubun, null, "Transaction_Complete", "", false, false);

                    }
                    catch (Exception ex)
                    {
                        setLogHandler(Gubun, ex, querySource, queryTransactionUpdate, false, true);
                    }
                    finally
                    {
                        cnt--;    
                    }
                    
                }
                
            }
            catch (Exception ex)
            {
                setLogHandler(Gubun, ex, querySource, queryTransaction, false, true);
                setLogHandler(Gubun, ex, querySource, queryTransactionUpdate, false, true);
            }

        }

        /// <summary>
        /// 프로그램 시작 전 Connection 정보 확인 및 화면에 보여줄 기준정보를 셋팅합니다.
        /// </summary>
        private bool SetBaseInfo()
        {
            //Connection 조건확인
            if (txtERPDB.Text == "Disconnection")
            {
                MessageBox.Show("ERP Connection이 연결되지 않았습니다.");
                return false;
            }
            else if (txtMESDB.Text == "Disconnection")
            {
                MessageBox.Show("MES Connection이 연결되지 않았습니다.");
                return false;
            }

            if (cbMailCheck.Text.Equals(""))
            {
                MessageBox.Show("오류시 메일발송 기준정보를 설정하세요.");
                return false;
            }
            else if (cbSMSCheck.Text.Equals(""))
            {
                MessageBox.Show("오류시 SMS발송 기준정보를 설정하세요.");
                return false;
            }
            else if (cbTimer.Text.Equals(""))
            {
                MessageBox.Show("프로그램 실행간격 기준정보를 설정하세요.");
                return false;
            }

            //ini파일 파싱 및 listBox 셋팅
            SetFileStream();

            //화면의 기준정보를 잠금설정 (최초 설정 후 Start할 경우 변경할 수 없음)
            setControlLock("S");

            return true;
        }


        /// <summary>
        /// 화면에 인터페이스 중인 테이블을 보여주기 위한 list를 Control 합니다.
        /// </summary>
        /// <param name="arrInfo">파싱하여 화면에 보여 줄 ini파일의 정보가 담겨있는 배열입니다.</param>
        /// <param name="Gubun">처리를 확인하는 구분자 입니다.</param>
        private void SetlistControl(string[] arrInfo, string Gubun)
        {
            int cnt = 0;

            foreach (var info in arrInfo)
            {
                cnt++;
                ListViewItem lstVi = new ListViewItem(cnt.ToString());
                lstVi.SubItems.Add(info.ToString());

                if (Gubun.ToUpper() == "ERP")
                    listERP.Items.Add(lstVi);
                else
                    listMES.Items.Add(lstVi);
            }
        }

        /// <summary>
        /// 화면에 있는 컨트롤들을 구분에 따라 잠금설정 or 해지설정을 합니다.
        /// </summary>
        /// <param name="Gubun">처리를 확인하는 구분자 입니다 (S: Start, E: End)</param>
        private void setControlLock(string Gubun)
        {
            //프로그램 시작점으로 컨트롤을 잠금
            if (Gubun.ToUpper() == "S")
            {
                btnMESChk.Enabled = false;
                btnERPChk.Enabled = false;
                btnStart.Enabled = false;
                cbMailCheck.Enabled = false;
                cbSMSCheck.Enabled = false;
                cbTimer.Enabled = false;

            }
            //프로그램 종료점으로 컨트롤을 잠금해지
            else
            {
                btnMESChk.Enabled = true;
                btnERPChk.Enabled = true;
                btnStart.Enabled = true;
                cbMailCheck.Enabled = true;
                cbSMSCheck.Enabled = true;
                cbTimer.Enabled = true;
            }
        }

        /// <summary>
        /// 로그 출력 및 메시지출력을 통합적으로 처리하는 메소드
        /// </summary>
        /// <param name="type">메시지의 타입입니다.</param>
        /// <param name="ex">예외 객체입니다.</param>
        /// <param name="querySource">조회하는 쿼리소스 입니다.</param>
        /// <param name="queryTarget ">처리해야 할 트랜젝션 쿼리입니다.</param>
        /// <param name="msgBoxChk">바탕화면에 메시지를 출력할 것인지 말 것인지에 대한 여부입니다.</param>
        /// <param name="error">애러여부를 확인하여 로그에 출력 해줄 것인지 판단한다.</param>
        private void setLogHandler(string type, Exception ex, string querySource, string queryTarget, bool msgBoxChk, bool error)
        {
            //error유무를 판별하여 로그에 출력할 때 표시
            if (error == true)
            {
                logCtrl.IOFileWrite(type + "[ERROR]", ex, querySource, queryTarget);
                
                //애러메일 송부내용
                _errorThrow.Add(string.Format("TYPE = {0} {1},  CONTENTS = {2} {3}", type, Environment.NewLine, ex.Message.ToString(), Environment.NewLine));
            }    
            else
                logCtrl.IOFileWrite(type, ex, querySource, queryTarget);

            //화면에 메시지를 출력해야 할 경우
            if (msgBoxChk == true)
                MessageBox.Show(ex.Message.ToString());
        }

        /// <summary>
        /// 주기적으로 호출되는 타이머 메소드이다. (프로그램의 시작점이기도 함)
        /// </summary>
        /// <param name="sender">타이머의 오브젝트 입니다.</param>
        /// <param name="e">타이머의 이벤트 입니다.</param>
        private void mainTimer_Tick(object sender, EventArgs e)
        {
            _errorThrow.Clear(); //매 주기마다 초기화 밑 애러가 발생될 때마다 기록을 남겨 로직이 끝나기 전 알람을 보낸다.


            mainTimer.Enabled = false; //프로세스가 실행되기 전 사용안함 처리, 그래야 프로세스가 실행되는 시간을 타이머가 계산하지 않음

            //Delay(3000);
            processStart();

            mainTimer.Enabled = true;   //프로세스가 완료된 후 다시 타이머를 적용
                
            //애러가 있으면 메일 및 SMS 전송
            if (_errorThrow.Count >= 1)
                setSend();

        }


        /// <summary>
        /// 데이터를 I/F하고 정상처리 되었다면 Flag를 변경해주는 메소드.
        /// </summary>
        /// <param name="hsInfo">ini파일의 정보가 담겨있는 HashTable 입니다.(각종 쿼리 및 정보)</param>
        /// <param name="tr">처리를 해야하는 데이터로우 입니다. (이 정보들이 조합되어 트랜젝션 쿼리가 생성 됩니다.)</param>
        /// <param name="Gubun">서버마다 시간을 표시하는 함수가 다르므로 구분할 수 있도록..</param>
        private string getUpdtQuery(Hashtable hsInfo, DataRow tr, string Gubun)
        {
            StringBuilder tempQuery = new StringBuilder();
            string tempTableName = hsInfo["SOURCE"].ToString();
            string tempSourceOption = hsInfo["SOURCE_OPTION"].ToString();
            string tempDate = string.Empty;
            string tempColumns = string.Empty;
            string tempValues = string.Empty;


            if (Gubun.ToUpper() == "MES")
                tempDate = "SYSDATE";
            else
                tempDate = "GETDATE()";

            //업데이트 할 테이블이름 추출(SELECT 한 쿼리의 테이블, 즉 Source쿼리)
            tempTableName = tempTableName.Substring(tempTableName.IndexOf("FROM") + 4, (tempTableName.IndexOf("WHERE") - tempTableName.IndexOf("FROM") - 4)).Trim();

            //업데이트 쿼리 조합
            tempQuery.Append("");
            tempQuery.Append(string.Format("UPDATE {0} SET ", tempTableName));

            if (tempTableName.ToUpper() == "T_IF_RCV_PROD_ORD_KO441")   //해당 테이블은 나중에 만들어졌기에 컬럼 이름이 일치하지 않음, 그러기에 프로그램에서 별도로 관리 (해당 테이블은 ERP_APPLY_FLAG1, 다른 테이블은 ERP_APPLY_FLAG)
                tempQuery.Append(string.Format("UPDT_USER_ID = 'INTERFACE', UPDT_DT =  {0}, ERP_APPLY_FLAG1 = 'Y' ", tempDate));
            else
                tempQuery.Append(string.Format("UPDT_USER_ID = 'INTERFACE', UPDT_DT =  {0}, {1} = {0}, {2} = 'Y' ", tempDate, (Gubun == "MES") ? "ERP_RECEIVE_DT" : "MES_RECEIVE_DT", (Gubun == "MES") ? "ERP_RECEIVE_FLAG" : "MES_RECEIVE_FLAG"));
            
            tempQuery.Append("WHERE 1=1 ");

            /*
            //이전 WHERE조건, INI파일에 TARGET_KEYS라는 MSSQL테이블의 키 값을 기입해뒀으나 실제 PK관리가 되어있는지 확인을 할 수 없음, 그러므로 KEY값만 가지고 UPDATE 조건을 사용하기에는
            //무리가 있다고 판단되어 주석처리 후 아래와 같은 전체컬럼 값으로 WHERE조건을 구성
            foreach (var hstSplit in hsInfo["TARGET_KEYS"].ToString().Split(','))
            {
                tempColumns = hsInfo[hstSplit].ToString().ToUpper().Trim();
                tempValues = tr[hstSplit].ToString();

                //조합된 컬럼명, 값을 쿼리에 셋팅
                tempQuery.Append(string.Format("AND {0} = {1} ", tempColumns, tempValues));
            }
            */

            //WHERE 조건 조합
            for (int cnt = 0; cnt < tr.Table.Columns.Count; cnt++)
            {
                //값이 있는 컬럼만 or 값이 SOURCE_OPTION이 맞는 것만
                if (tr[cnt].ToString().Length > 0 && tr[cnt].GetType().Name.ToUpper() == tempSourceOption.ToUpper())
                { 
                    tempColumns = tr.Table.Columns[cnt].ColumnName.ToString();
                    tempValues  = string.Format("'{0}'", tr[cnt].ToString());

                    //조합된 컬럼명, 값을 쿼리에 셋팅
                    tempQuery.Append(string.Format("AND {0} = {1} ", tempColumns, tempValues));    
                }
            }

            return tempQuery.ToString();
        }
        
        
        ////딜레이 함수 (스레드를 죽이면 프로세스가 멈추므로 딜레이 시간을 주고 멈출 수 있는 함수)
        //private static DateTime Delay(int MS)
        //{
        //    DateTime ThisMoment = DateTime.Now;
        //    TimeSpan duration = new TimeSpan(0, 0, 0, 0, MS);
        //    DateTime AfterWards = ThisMoment.Add(duration);

        //    while (AfterWards >= ThisMoment)
        //    {
        //        System.Windows.Forms.Application.DoEvents();
        //        ThisMoment = DateTime.Now;
        //    }

        //    return DateTime.Now;
        //}


        /// <summary>
        /// ini파일을 파싱하여 생성된 해쉬테이블을 기준으로 트랜잭션 쿼리를 조합한다.
        /// </summary>
        /// <param name="hsInfo">ini파일의 정보가 담겨있는 HashTable 입니다.(각종 쿼리 및 정보)</param>
        /// <param name="tr">처리를 해야하는 데이터로우 입니다. (이 정보들이 조합되어 트랜젝션 쿼리가 생성 됩니다.)</param>
        /// <param name="Gubun">서버마다 시간을 표시하는 함수가 다르므로 구분할 수 있도록..</param>
        /// <param name="option">쿼리의 실행타입을 확인(insert, update)</param>
        private string getQuery(Hashtable hsInfo, DataRow tr, string Gubun, string option)
        {
            string tempHeaderQuery = string.Empty;                      //최종 조합 될 쿼리 (해당 변수를 리턴 함)
            string tempColumnQuery = string.Empty;                      //최종 조합 될 쿼리 (해당 변수를 리턴 함)
            string tempValueQuery = string.Empty;                       //최종 조합 될 VALUSE 쿼리
            string tempColumn_DT = string.Empty;
            string tempColumn_Flag = string.Empty;

            switch(option.ToUpper())
            {
                case "INSERT" :
                    tempHeaderQuery = string.Format("{0} (", hsInfo["TARGET"].ToString());
                    tempColumnQuery = "";
                    tempColumn_DT = (Gubun.ToUpper() == "MES") ? "ERP_RECEIVE_DT" : "MES_RECEIVE_DT";
                    tempColumn_Flag = (Gubun.ToUpper() == "MES") ? "ERP_RECEIVE_FLAG" : "MES_RECEIVE_FLAG";

                    for (int cnt = 0; cnt < tr.Table.Columns.Count; cnt++)
                    {
                        if (cnt != tr.Table.Columns.Count - 1)
                        {
                            //쿼리의 insert 컬럼 셋팅
                            tempColumnQuery += tr.Table.Columns[cnt].ColumnName.ToString() + ",";

                            //컬럼이름이 RECEIVE_DT일 경우는 현재 시간이 들어가야 함, 한꺼번에 트랜잭션이 발생하는 로직이기에 MES테이블에 먼저 UPDATE를 할 수 없기에 INSERT 할 때 처리
                            if (tr.Table.Columns[cnt].ColumnName.ToString().ToUpper().Trim() == tempColumn_DT.Trim())
                            {
                                if (Gubun.ToUpper() == "MES")
                                    tempValueQuery += "GETDATE(), ";
                                else
                                    tempValueQuery += "SYSDATE, ";
                            }
                            //컬럼이름이 RECEIVE_FLAG일 경우는 실제 INSERT 하는 중이니 'Y'처리 해야 함
                            else if (tr.Table.Columns[cnt].ColumnName.ToString().ToUpper().Trim() == tempColumn_Flag.Trim()) 
                            {
                                tempValueQuery += "'Y', ";
                            }
                            //RECEIVE_FLAG, DT가 아닌 나머지는 정상적으로 처리
                            else
                            {
                                switch (tr[cnt].GetType().Name.ToUpper()) //데이터의 형식을 확인 함
                                {
                                    case "DATETIME":   //데이터형식이 시간                                    
                                    
                                        //MSSQL와 ORACLE의 이기종간 데이터치환이 다르기 때문에 쿼리도 다르게 만듬
                                        if(Gubun.ToUpper() == "MES")
                                            tempValueQuery += string.Format("'{0}', ", Convert.ToDateTime(tr[cnt]).ToString("yyyy.MM.dd HH:mm:ss"));
                                        else
                                            tempValueQuery += string.Format("TO_DATE('{0}' , 'YYYYMMDDHH24MISS'), ", Convert.ToDateTime(tr[cnt]).ToString("yyyyMMddHHmmss"));

                                        break;

                                    case "DECIMAL" :    //데이터형식이 숫자
                                        tempValueQuery += string.Format("{0}, ", tr[cnt].ToString()).Trim();
                                        break;

                                    case "DBNULL" : //데이터형식이 null
                                        tempValueQuery += "NULL, ";
                                        break;

                                    default :       //나머지(string)
                                        tempValueQuery += string.Format("'{0}', ", tr[cnt].ToString().Trim());
                                        break;
                                }
                            }
                        }
                        //쿼리의 마지막문은 콤마가 아닌 괄호로 마무리를 해야 하기 때문에
                        else   
                        {
                            tempColumnQuery += tr.Table.Columns[cnt].ColumnName.ToString() + ") VALUES(";

                            //컬럼이름이 RECEIVE_DT일 경우는 현재 시간이 들어가야 함, 한꺼번에 트랜잭션이 발생하는 로직이기에 MES테이블에 먼저 UPDATE를 할 수 없기에 INSERT 할 때 처리
                            if (tr.Table.Columns[cnt].ColumnName.ToString().ToUpper().Trim() == tempColumn_DT.Trim())
                            {
                                if (Gubun.ToUpper() == "MES")
                                    tempHeaderQuery = tempHeaderQuery + tempColumnQuery + tempValueQuery + "GETDATE())";
                                else
                                    tempHeaderQuery = tempHeaderQuery + tempColumnQuery + tempValueQuery + "SYSDATE)";
                            }
                            //컬럼이름이 RECEIVE_FLAG일 경우는 실제 INSERT 하는 중이니 'Y'처리 해야 함
                            else if (tr.Table.Columns[cnt].ColumnName.ToString().ToUpper().Trim() == tempColumn_Flag.Trim())
                            {
                                tempHeaderQuery = tempHeaderQuery + tempColumnQuery + tempValueQuery + "'Y')";
                            }
                            //RECEIVE_FLAG, DT가 아닌 나머지는 정상적으로 처리
                            else
                            {
                                switch (tr[cnt].GetType().Name.ToUpper()) //데이터의 형식을 확인 함
                                {
                                    case "DATETIME":   //데이터형식이 시간                                    

                                        //MSSQL와 ORACLE의 이기종간 데이터치환이 다르기 때문에 쿼리도 다르게 만듬
                                        if (Gubun.ToUpper() == "MES")
                                            tempHeaderQuery = tempHeaderQuery + tempColumnQuery + tempValueQuery + string.Format("'{0}')", Convert.ToDateTime(tr[cnt]).ToString("yyyy.MM.dd HH:mm:ss"));
                                        else
                                            tempHeaderQuery = tempHeaderQuery + tempColumnQuery + tempValueQuery + string.Format("TO_DATE('{0}', 'YYYYMMDDHH24MISS'))", Convert.ToDateTime(tr[cnt]).ToString("yyyyMMddHHmmss"));

                                        break;

                                    case "DECIMAL":    //데이터형식이 숫자
                                        tempHeaderQuery = tempHeaderQuery + tempColumnQuery + tempValueQuery + string.Format("{0})", tr[cnt].ToString().Trim());
                                        break;

                                    case "DBNULL": //데이터형식이 null
                                        tempHeaderQuery = tempHeaderQuery + tempColumnQuery + tempValueQuery + "NULL)";
                                        break;

                                    default:       //나머지(string)
                                        tempHeaderQuery = tempHeaderQuery + tempColumnQuery + tempValueQuery + string.Format("'{0}')", tr[cnt].ToString().Trim());
                                        break;
                                }
                            }
                        }
                    }
                        
                    break;

                case "UPDATE" :
                    tempHeaderQuery = hsInfo["TARGET"].ToString() + " WHERE 1=1 ";
                    tempColumnQuery = "";

                    foreach (var hstSplit in hsInfo["TARGET_KEYS"].ToString().Split(','))
                        tempColumnQuery += string.Format("AND {0} = '{1}'", hstSplit.ToString().Trim(), tr[hstSplit].ToString().Trim());

                    tempHeaderQuery = tempHeaderQuery + tempColumnQuery;

                    break;

            }

            return tempHeaderQuery;
        }

        /// <summary>
        /// 기준정보를 사용하며 실제 ini의 파일에 내용을 확인하여 SMS 및 Mail을 전송하는 메소드를 호출한다. 
        /// </summary>
        private void setSend()
        {
            Hashtable tempHs = new Hashtable();
            Hashtable smsHs = new Hashtable();
            Hashtable mailHs = new Hashtable();
            string tempIniKey = string.Empty;
            string Gubun = "SET";
            string alarmTime = string.Empty;
            string smsFlag = "N";
            string mailFlag = "N";

            for (int index = 0; index < _iniSetInfo.Count(); index++)
            {
                for (int index2 = 0; index2 < _iniSetInfo[index].Count(); index2++)
                {
                    tempIniKey = _iniSetInfo[index][index2].Substring(0, _iniSetInfo[index][index2].IndexOf("="));
                    SethashTable(tempHs, tempIniKey, index, index2, Gubun);
                }

                //루프를 돌기 때문에 순서와 영향이 있을 수 있다. 
                //그러기에 조건이 만족할 경우 변수에 플래그처리를 한 뒤 루프를 끝낸 뒤 후속처리를 진행한다.

                //ini파일에 메일 및 SMS보낼 시간을 변수에 담는다.
                if (tempHs["SOURCE"].ToString() == "ALARM_TIME")
                    alarmTime = tempHs["TARGET"].ToString();

                //sms을 보내겠다고 설정되어 있으며 hst테이블의 값 중 PHONE이 있을 경우
                if (cbSMSCheck.Text == "Y" && tempHs["SOURCE"].ToString() == "PHONE_NUMBER" && tempHs["TARGET"].ToString().Length > 0)
                { 
                    smsFlag = "Y";
                    smsHs = (Hashtable)tempHs.Clone(); //sms 전용 해쉬테이블로 복사 (루프에서 처리하지 않기에 별도의 해쉬테이블에 저장), 일반적인 복사(A = B)를 할 경우 얕은 복사로 B의 데이터를 CLARE할 경우 A의 데이터도 없어지기에 복제
                }      
                        
                //mail을 보내겠다고 설정되어 있으며 hst테이블의 값이 MAIL이 있을 경우
                if (cbMailCheck.Text == "Y" && tempHs["SOURCE"].ToString() == "EMAIL_ADRESS" && tempHs["TARGET"].ToString().Length > 0)
                {
                    mailFlag = "Y";
                    mailHs = (Hashtable)tempHs.Clone();    //mail 전용 해쉬테이블로 복사 (루프에서 처리하지 않기에 별도의 해쉬테이블에 저장), 일반적인 복사(A = B)를 할 경우 얕은 복사로 B의 데이터를 CLARE할 경우 A의 데이터도 없어지기에 복제
                }

                tempHs.Clear();
            }


            //시스템의 현재시간을 확인하여 기준정보에 입력되어 있었을 경우 메소드를 호출한다.
            //야간에는 알람을 받지 않도록 하기 위해서 실행되는 시간(HH)을 확인하여 알람전송
            if (alarmTime.IndexOf(DateTime.Now.ToString("HH").ToString()) != -1)
            {
                try
                {
                    //상단 프로세스에서 조건이 맞아 Flag가 'Y'로 변경되었을 경우
                    if (smsFlag == "Y")
                        comFunc.setSendSMS(_errorThrow, smsHs);

                    //상단 프로세스에서 조건이 맞아 Flag가 'Y'로 변경되었을 경우
                    if (mailFlag == "Y")
                        comFunc.setSendMail(_errorThrow, mailHs);
                }
                catch (Exception ex)
                {
                    setLogHandler("알람 중 오류 발생", ex, "", "", false, true);
                }

            }       
        }

    }

}



