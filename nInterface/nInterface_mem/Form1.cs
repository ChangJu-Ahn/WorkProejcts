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
using System.Diagnostics;   //for verify process
using System.Timers;

namespace nInterface
{
    public partial class Form1 : Form
    {
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
        /// It is string list object to bring on the basesis ini file (information of ERP, MES and user Setting)
        /// </summary>
        List<List<string>> _iniERPInfo = new List<List<string>>();
        List<List<string>> _iniMESInfo = new List<List<string>>();
        List<List<string>> _iniSetInfo = new List<List<string>>();

        /// <summary>
        /// It is string list object to substitute addition Error message.
        /// </summary>
        List<string> _errorThrow = new List<string>();
        #endregion

        #region It Generate an Database object that is responsible for connection (ERP, MES)
        /// <summary>
        /// Object of database(MES, ERP)
        /// </summary>
        static nInterface.DBConn _DBConn = new DBConn();
        #endregion

        #region Other variables.
        /// <summary>
        /// It is Member variable to use Program.
        /// </summary>
        static iniFileHandle.Win32Reg _Win32Reg = null;
        static nInterface.LogControl logCtrl = new LogControl(_LogPath);
        static nInterface.CommonFunction comFunc = new CommonFunction();
        string _smsFlag, _mailFlag;
        int repetTime;
        int timerBaseTime = 1000;
        #endregion

        #region variables for use Thread.
        /// <summary>
        /// It is Member variable of multi Thread.
        /// </summary>
        System.Timers.Timer mainTimer = new System.Timers.Timer();
        delegate void DelTimerText(string data);
        private volatile bool isStop = false;
        #endregion


        /// <summary>
        /// it is the program initial entry point
        /// </summary>
        public Form1()
        {
            processCheck();
            InitializeComponent();
            notifyIcon1.Visible = false; //트레이 아이콘을 숨긴다.            
        }


        /// <summary>
        /// 실행중인 프로세스가 있는지 확인합니다.
        /// 만약 프로세스가 실행 중이라면 더 이상의 프로그램을 실행하지 않습니다.
        /// </summary>
        private void processCheck()
        {
            Process[] processList = Process.GetProcessesByName("nInterface_mem");   //this is Program process name

            if (processList.Length > 1)
            {
                MessageBox.Show("이미 프로세스가 실행 중 입니다.");
                programEnd();
            }

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
            mainTimer.Interval = timerBaseTime;
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
                _Win32Reg = new iniFileHandle.Win32Reg(_initPath + "ERP-MES.ini");
                iniSection = _Win32Reg.GetSectionNames();
                SetiniInfo(iniSection, "ERP");
                SetlistControl(iniSection, "ERP");     //listView 셋팅 (배열이 초기화 되기 전 셋팅)
            }

            if (_iniMESInfo.Count == 0)
            {
                _Win32Reg.szFileName = _initPath + "MES-ERP.ini";
                //Win32Reg = new iniFileHandle.Win32Reg(_initPath + "MES-ERP.ini");
                iniSection = _Win32Reg.GetSectionNames();
                SetiniInfo(iniSection, "MES");
                SetlistControl(iniSection, "MES");     //listView 셋팅 (배열이 초기화 되기 전 셋팅)
            }

            if (_iniSetInfo.Count == 0)
            {
                _Win32Reg = new iniFileHandle.Win32Reg(_setPath + "setInfo.ini");
                iniSection = _Win32Reg.GetSectionNames();
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


                tempArray = _Win32Reg.GetPairsBySection(iniInfo[index]);

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
        /// 실행 프로세스를 멀티스레드로 관리합니다. 
        /// 이때 타이머 반복시간, sms 및 mail Flag를 전역변수에 대입합니다.
        /// </summary>
        private void TimerThreadStart()
        {
            repetTime = Convert.ToInt32(cbTimer.Text) * 60000;
            _smsFlag = cbSMSCheck.Text;
            _mailFlag = cbMailCheck.Text;

            //타이머시간 셋팅(밀리초이기 때문에 분을 만들기 위해서 화면에서 전달받은 숫자 * 60초를 해야 함 (1초 = 1000초))
            mainTimer.Elapsed += new ElapsedEventHandler(mainTimer_Elapsed);
            mainTimer.Start();
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

                TimerThreadStart();
            }
            catch (Exception ex)
            {
                setLogHandler("Start_Click..", ex, "", "", true, true);
            }
        }

        //델리케이트용 함수
        #region
        /// <summary>
        /// UI스레드에 값을 대입하기 위한 델리게이트용 함수입니다.
        /// </summary>
        private void setStartTimerText(string data)
        {
            if (lblbeginTime.InvokeRequired)
            {
                DelTimerText call = new DelTimerText(setStartTimerText);
                this.Invoke(call, data);
            }
            else
                lblbeginTime.Text = data;
        }

        private void setEndTimerText(string data)
        {
            if (lblEndTime.InvokeRequired)
            {
                DelTimerText call = new DelTimerText(setEndTimerText);
                this.Invoke(call, data);
            }
            else
                lblEndTime.Text = data;
        }
        #endregion


        /// <summary>
        /// begin Program Process.
        /// </summary>
        private void processStart()
        {
            //타이머를 실행할 때마다 최종시간을 화면에 표시
            setStartTimerText(System.DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss"));

            //mes interface 로직 실행
            mesDataProcess();

            //erp interface 로직 실행
            erpDataProcess();

            //타이머가 종료될 때마다 최종시간을 화면에 표시
            setEndTimerText(System.DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss"));
        }


        /// <summary>
        /// If you click the Stop button it is the initial entry point.
        /// </summary>
        private void Stop_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("프로그램을 종료하시겠습니까?", "종료확인", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (isStop == false && mainTimer.Enabled == false && Convert.ToInt32(mainTimer.Interval) != timerBaseTime)
                {
                    isStop = true;
                    btnEnd.Enabled = false;
                    MessageBox.Show("프로세스가 실행 중으로 종료예약을 설정합니다.");
                }
                else
                {
                    //DB연결 종료
                    if (!_DBConn.ConnectionERPString.Equals(""))
                        _DBConn.CloseERPDBConn();

                    if (!_DBConn.ConnectionMESString.Equals(""))
                        _DBConn.CloseMESDBConn();

                    programEnd();
                }
            }
        }


        /// <summary>
        /// 멀티스레드를 종료하고 최종 프로그램을 종료합니다.
        /// </summary>
        private void programEnd()
        {
            notifyIcon1.Visible = false;
            mainTimer.Stop();
            mainTimer.Close();

            Application.ExitThread();
            Environment.Exit(0);
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
                    //반복문 쿼리변수 및 상태변수 초기화
                    queryTransaction = "";
                    queryTransactionUpdate = "";
                    transactionCheck = false;

                    //실제 트랜젝션이 일어날 때 오류가 발생되었어도 해당 쿼리는 건너뛰고 다음쿼리를 진행하기 위해 중첩 try-catch 사용 
                    //첫 번째 try-catch의 용도 : SELECT문이 오류가 있어도 다음 쿼리를 실행하기 위한 예외처리
                    //두 번째 try-catch의 용도 : SELECT된 데이터를 트랜젝션 쿼리가 발생될 때 문제부분을 건너뛰고 실행하기 위함
                    try
                    {
                        queryTransaction = getQuery(hst, dtResult.Rows[cnt - 1], Gubun.ToUpper(), targetOption);

                        //update문 같은경우는 mes와 erp에 동일한 정보에 업데이트를 해줘야 함, 그러므로 update 로직일 때는 기존 쿼리를 업데이트문 쿼리에 대입
                        if (targetOption == "INSERT")
                        {
                            queryTransactionUpdate = getUpdtQuery(hst, dtResult.Rows[cnt - 1], Gubun.ToUpper());
                        }
                        else if (targetOption == "UPDATE")
                        {
                            if (hst["BOTH_TRANSACTION"].ToString().Trim().ToUpper().Equals("Y"))
                                queryTransactionUpdate = queryTransaction;
                            else if (hst["BOTH_TRANSACTION"].ToString().Trim().ToUpper().Equals("N"))
                                queryTransactionUpdate = "SINGLE_UPDATE";
                            else
                                queryTransactionUpdate = "UNUSUAL_ROUTE";
                        }
                        else
                        {
                            setLogHandler(Gubun, null, "Unusual_Route, Excute break constraction", "", false, false);
                            break;
                        }

                        //is written transaction log and state log before executing transaction query.
                        setLogHandler(Gubun, null, "Transaction_Start(" + targetOption + ")", queryTransaction, false, false);
                        setLogHandler(Gubun, null, "Transaction_Start(" + targetOption + ")", queryTransactionUpdate, false, false);

                        // handling process of really data 
                        if (queryTransactionUpdate.Equals("UNUSUAL_ROUTE"))
                        {
                            setLogHandler(Gubun, null, "Unusual_Route, Excute break constraction", "", false, false);
                            break;
                        }
                        else if (queryTransactionUpdate.Equals("SINGLE_UPDATE"))
                        {
                            transactionCheck = _DBConn.ExecuteNonQuery(queryTransaction, Gubun.ToUpper());
                        }
                        else
                        {
                            transactionCheck = _DBConn.ExecuteTransactionNonQuery(queryTransaction, queryTransactionUpdate, Gubun.ToUpper());
                        }

                        //If it return normality value, is written completion log.
                        if (transactionCheck == true)
                            setLogHandler(Gubun, null, "Transaction_Completion(" + targetOption + ")", "", false, false);
                        else
                            setLogHandler(Gubun, null, "Transaction_Failure(" + targetOption + ")", "", false, false);

                    }
                    catch (Exception ex)
                    {
                        setLogHandler(Gubun, ex, querySource, queryTransaction, false, true);
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
                setLogHandler(Gubun, ex, querySource, "", false, true);
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
        private void mainTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (Convert.ToInt32(mainTimer.Interval) == timerBaseTime) mainTimer.Interval = repetTime;
            if (isStop == true) programEnd();

            _errorThrow.Clear(); //매 주기마다 초기화 밑 애러가 발생될 때마다 기록을 남겨 로직이 끝나기 전 알람을 보낸다.

            mainTimer.Enabled = false; //프로세스가 실행되기 전 사용안함 처리, 그래야 프로세스가 실행되는 시간을 타이머가 계산하지 않음
            DBConnection_Check(1);     //DB Open

            //Delay(3000);
            processStart();

            //애러가 있으면 메일 및 SMS 전송
            if (_errorThrow.Count >= 1)
                setSend();

            DBConnection_Check(2);      //DB Close
            mainTimer.Enabled = true;   //프로세스가 완료된 후 다시 타이머를 적용
            if (isStop == true) programEnd();
        }


        /// <summary>
        /// DB연결을 체크하여 Open 또는 Close를 실시한다.
        /// </summary>
        /// <param name="separator">DB를 Open 또는 Close를 할 수 있는 구분자입니다(1: Opne, 2: Close)</param>
        private void DBConnection_Check(int separator)
        {
            if (separator == 1)
            {
                _DBConn.OpenERPDBConn();
                _DBConn.OpenMESDBConn();
            }
            else
            {
                //모든 실행단 이후에는 db를 close 한다.
                _DBConn.CloseERPDBConn();
                _DBConn.CloseMESDBConn();
            }
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
            string tempDate = string.Empty;

            if (Gubun.ToUpper() == "MES")
                tempDate = "SYSDATE";
            else
                tempDate = "GETDATE()";

            //업데이트 할 테이블이름 추출(SELECT 한 쿼리의 테이블, 즉 Source쿼리)
            tempTableName = tempTableName.Substring(tempTableName.IndexOf("T_IF"), tempTableName.IndexOf("WHERE") - tempTableName.IndexOf("T_IF")).ToString().Trim().ToUpper();

            //업데이트 쿼리 조합
            tempQuery.Append("");
            tempQuery.Append(string.Format("UPDATE {0} SET ", tempTableName));

            if (tempTableName.ToUpper() == "T_IF_RCV_PROD_ORD_KO441")   //해당 테이블은 나중에 만들어졌기에 컬럼 이름이 일치하지 않음, 그러기에 프로그램에서 별도로 관리 (해당 테이블은 ERP_APPLY_FLAG1, 다른 테이블은 ERP_APPLY_FLAG)
                tempQuery.Append(string.Format("UPDT_USER_ID = 'INTERFACE', UPDT_DT =  {0}, ERP_APPLY_FLAG1 = 'Y' ", tempDate));
            else
                tempQuery.Append(string.Format("UPDT_USER_ID = 'INTERFACE', UPDT_DT =  {0}, {1} = {0}, {2} = 'Y' ", tempDate, (Gubun == "MES") ? "ERP_RECEIVE_DT" : "MES_RECEIVE_DT", (Gubun == "MES") ? "ERP_RECEIVE_FLAG" : "MES_RECEIVE_FLAG"));

            tempQuery.Append("WHERE 1=1 ");

            /*bottom code is update query(bottom part of the update query "WHERE")*/
            foreach (var hstSplit in hsInfo["TARGET_KEYS"].ToString().Split(','))
            {
                string tempColumnName = hstSplit.ToString().ToUpper().Trim().ToUpper();
                string tempColumnType = tr[tempColumnName].GetType().Name.ToString().Trim().ToUpper();
                object tempColumnContent = tr[tempColumnName];

                tempQuery.Append(
                        string.Format(" AND {0} = {1}", tempColumnName, GetSwitchQueryText(tempColumnType, tempColumnContent, Gubun.ToUpper(), true))
                    );
            }

            return tempQuery.ToString();
        }

        /// <summary>
        /// This method return a query using a switch statement.
        /// </summary>
        /// <param name="columnType">The type of Column</param>
        /// <param name="columnContent">The object that have real value of Column</param>
        /// <param name="guBunFlg">The identifier of database(ERP or MES)</param>
        /// <param name="infoReverseFlg">The identifier that decide to information direction</param>
        private string GetSwitchQueryText(string columnType, object columnContent, string guBunFlg, bool infoReverseFlg)
        {
            string tempSwitchQuery = string.Empty;

            switch (columnType)
            {
                case "DATETIME":
                    if (infoReverseFlg == true)
                    {
                        if (guBunFlg.ToUpper() == "MES")
                            tempSwitchQuery = string.Format("TO_DATE({0}, 'YYYYMMDDHH24MISS')", Convert.ToDateTime(columnContent).ToString("yyyyMMddHHmmss"));
                        else
                            tempSwitchQuery = string.Format("'{0}'", Convert.ToDateTime(columnContent).ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    }
                    else
                    {
                        if (guBunFlg.ToUpper() == "ERP")
                            tempSwitchQuery = string.Format("TO_DATE({0}, 'YYYYMMDDHH24MISS')", Convert.ToDateTime(columnContent).ToString("yyyyMMddHHmmss"));
                        else
                            tempSwitchQuery = string.Format("'{0}'", Convert.ToDateTime(columnContent).ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    }

                    break;

                case "DECIMAL":    //데이터형식이 숫자
                    tempSwitchQuery = string.Format("{0}", columnContent.ToString().Trim());
                    break;

                case "DBNULL": //데이터형식이 null
                    tempSwitchQuery = string.Format("NULL");
                    break;

                default:       //나머지(string)
                    tempSwitchQuery = string.Format("'{0}'", columnContent.ToString().Trim().Replace("'", "''"));    //2016.11.14, ahncj : If A among various values entered,
                    break;
            }

            return tempSwitchQuery;
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
            StringBuilder tempValueQuery = new StringBuilder();
            StringBuilder tempHeaderQuery = new StringBuilder();
            string tempColumnName = string.Empty;
            string tempColumnType = string.Empty;
            object tempColumnContent;

            switch (option.ToUpper())
            {
                case "INSERT":
                    string tempColumn_DT = (Gubun.ToUpper() == "MES") ? "ERP_RECEIVE_DT" : "MES_RECEIVE_DT";
                    string tempColumn_Flag = (Gubun.ToUpper() == "MES") ? "ERP_RECEIVE_FLAG" : "MES_RECEIVE_FLAG";
                    tempHeaderQuery.Append(string.Format("{0} (", hsInfo["TARGET"].ToString()));

                    for (int cnt = 0; cnt < tr.Table.Columns.Count; cnt++)
                    {
                        tempColumnName = tr.Table.Columns[cnt].ColumnName.ToString().ToUpper().Trim();          //it is columns of target that is changed(to column)
                        tempColumnType = tr[tempColumnName].GetType().Name.ToString().ToUpper().Trim();                   //this is the previous column type.
                        tempColumnContent = tr[tempColumnName];                                                 //this is the obejct that contains the value of database.

                        if (cnt != tr.Table.Columns.Count - 1)
                        {
                            //쿼리의 insert 컬럼 셋팅
                            tempHeaderQuery.Append(tempColumnName + ",");

                            //컬럼이름이 RECEIVE_DT일 경우는 현재 시간이 들어가야 함, 한꺼번에 트랜잭션이 발생하는 로직이기에 MES테이블에 먼저 UPDATE를 할 수 없기에 INSERT 할 때 처리
                            if (tempColumnName == tempColumn_DT)
                            {
                                if (Gubun.ToUpper() == "MES")
                                {
                                    tempValueQuery.Append("GETDATE(), ");
                                }
                                else
                                {
                                    tempValueQuery.Append("SYSDATE, ");
                                }
                            }
                            //컬럼이름이 RECEIVE_FLAG일 경우는 실제 INSERT 하는 중이니 'Y'처리 해야 함
                            else if (tempColumnName == tempColumn_Flag)
                            {
                                tempValueQuery.Append("'Y', ");
                            }
                            //RECEIVE_FLAG, DT가 아닌 나머지는 정상적으로 처리
                            else
                            {
                                tempValueQuery.Append(string.Format("{0}, ", GetSwitchQueryText(tempColumnType, tempColumnContent, Gubun, false)));
                            }
                        }
                        //쿼리의 마지막문은 콤마가 아닌 괄호로 마무리를 해야 하기 때문에
                        else
                        {
                            tempHeaderQuery.Append(tempColumnName + ") VALUES(");

                            //컬럼이름이 RECEIVE_DT일 경우는 현재 시간이 들어가야 함, 한꺼번에 트랜잭션이 발생하는 로직이기에 MES테이블에 먼저 UPDATE를 할 수 없기에 INSERT 할 때 처리
                            if (tempColumnName == tempColumn_DT)
                            {
                                if (Gubun.ToUpper() == "MES")
                                {
                                    tempHeaderQuery.Append(tempValueQuery + "GETDATE())");
                                }
                                else
                                {
                                    tempHeaderQuery.Append(tempValueQuery + "SYSDATE)");
                                }
                            }
                            //컬럼이름이 RECEIVE_FLAG일 경우는 실제 INSERT 하는 중이니 'Y'처리 해야 함
                            else if (tempColumnName == tempColumn_Flag)
                            {
                                tempHeaderQuery.Append(tempValueQuery + "'Y')");
                            }
                            //RECEIVE_FLAG, DT가 아닌 나머지는 정상적으로 처리
                            else
                            {
                                tempHeaderQuery.Append(tempValueQuery + string.Format("{0})", GetSwitchQueryText(tempColumnType, tempColumnContent, Gubun, false)));
                            }
                        }
                    }

                    break;

                case "UPDATE":
                    tempHeaderQuery.Append(hsInfo["TARGET"].ToString());

                    //If updated all information
                    if (hsInfo["BOTH_TRANSACTION"].ToString() == "Y")
                    {
                        tempHeaderQuery.Append(" WHERE 1=1 ");

                        foreach (var hstSplit in hsInfo["TARGET_KEYS"].ToString().Split(','))
                        {
                            tempColumnName = hstSplit.ToString().Trim();
                            tempColumnContent = tr[tempColumnName];

                            tempHeaderQuery.Append(string.Format("AND {0} = '{1}' ", tempColumnName.Trim(), tempColumnContent.ToString().Trim().Replace("'", "''")));
                        }

                        break;
                    }
                    //If updated single information
                    else
                    {
                        int cnt = 0;

                        /*bottom code is update query(upper part of the update query "WHERE")*/
                        foreach (var hstSplit in hsInfo["TARGET_COLUMN"].ToString().Split(','))
                        {
                            tempColumnName = hstSplit.ToString().ToUpper().Trim();
                            tempColumnType = tr[tempColumnName].GetType().Name.ToString().ToUpper().Trim();                   //this is the previous column type.
                            tempColumnContent = tr[tempColumnName];

                            if (cnt == 0)
                            {
                                tempHeaderQuery.Append(string.Format(" {0} = {1} ", tempColumnName, GetSwitchQueryText(tempColumnType, tempColumnContent, Gubun, false)));
                                //tempHeaderQuery.Append(string.Format(" {0} = '{1}' ", tempColumnName.Trim(), tempColumnContent.ToString().Trim().Replace("'", "''")));
                            }
                            else
                            {
                                tempHeaderQuery.Append(string.Format(", {0} = {1} ", tempColumnName, GetSwitchQueryText(tempColumnType, tempColumnContent, Gubun, false)));
                                //tempHeaderQuery.Append(string.Format(" ,{0} = '{1}' ", tempColumnName.Trim(), tempColumnContent.ToString().Trim().Replace("'", "''")));
                            }

                            cnt++;
                        }

                        tempHeaderQuery.Append(" WHERE 1=1 ");

                        /*bottom code is update query(bottom part of the update query "WHERE")*/
                        foreach (var hstSplit in hsInfo["TARGET_KEYS"].ToString().Split(','))
                        {
                            tempColumnName = hstSplit.ToString().ToUpper();
                            tempColumnContent = tr[hstSplit.ToString().Trim()];

                            tempHeaderQuery.Append(string.Format("AND {0} = '{1}' ", tempColumnName.Trim(), tempColumnContent.ToString().Trim().Replace("'", "''")));
                        }

                        break;
                    }
            }

            return tempHeaderQuery.ToString();
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
                if (_smsFlag == "Y" && tempHs["SOURCE"].ToString() == "PHONE_NUMBER" && tempHs["TARGET"].ToString().Length > 0)
                {
                    smsFlag = "Y";
                    smsHs = (Hashtable)tempHs.Clone(); //sms 전용 해쉬테이블로 복사 (루프에서 처리하지 않기에 별도의 해쉬테이블에 저장), 일반적인 복사(A = B)를 할 경우 얕은 복사로 B의 데이터를 CLARE할 경우 A의 데이터도 없어지기에 복제
                }

                //mail을 보내겠다고 설정되어 있으며 hst테이블의 값이 MAIL이 있을 경우
                if (_mailFlag == "Y" && tempHs["SOURCE"].ToString() == "EMAIL_ADRESS" && tempHs["TARGET"].ToString().Length > 0)
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
                //상단 프로세스에서 조건이 맞아 Flag가 'Y'로 변경되었을 경우
                if (smsFlag == "Y")
                {
                    try
                    {
                        comFunc.setSendSMS(_errorThrow, smsHs);
                    }
                    catch (Exception ex)
                    {
                        setLogHandler("알람 중 오류 발생(SMS)", ex, "", "", false, true);
                    }
                }

                //상단 프로세스에서 조건이 맞아 Flag가 'Y'로 변경되었을 경우
                if (mailFlag == "Y")
                {
                    try
                    {
                        comFunc.setSendMail(_errorThrow, mailHs);
                    }
                    catch (Exception ex)
                    {
                        setLogHandler("알람 중 오류 발생(MAIL)", ex, "", "", false, true);
                    }
                }
            }
        }

        #region this is control program code of display part.
        /// <summary>
        ///It is function when doing NotifyIcon double click
        /// </summary>
        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            this.Visible = true; // 폼의 표시

            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal; // 최소화를 멈춘다 

            this.Activate(); // 폼을 활성화 시킨다
            this.notifyIcon1.Visible = false;

            this.Icon = new Icon(Environment.CurrentDirectory + @"\nInterface_Icon.ico"); //reset form1 icon
        }


        /// <summary>
        ///It is function when closing in form1
        /// </summary>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.ToTray();
        }


        /// <summary>
        ///This is a function that is used without being terminated when the close button is pressed.
        /// </summary>
        private void ToTray()
        {
            this.Hide();
            notifyIcon1.Visible = true;
            notifyIcon1.ContextMenuStrip = contextMenuStrip1;
            notifyIcon1.ShowBalloonTip(100, "nepes Interface Program", "This is nepes Interface Program(ERP<->MES)", ToolTipIcon.Info);
        }


        /// <summary>
        ///트레이 아이콘을 설정하기 위한 이벤트처리 함수입니다(메뉴바에서 Open버튼을 눌렀을 경우)
        /// </summary>
        private void contextMenuOpen_Cilck(object sender, EventArgs e)
        {
            notifyIcon_DoubleClick(sender, e);
        }


        /// <summary>
        ///트레이 아이콘을 설정하기 위한 이벤트처리 함수입니다.(메뉴바에서 Close버튼을 눌렀을 경우)
        /// </summary>
        private void contextMenuClose_Cilck(object sender, EventArgs e)
        {
            Stop_Click(sender, e);
        }
        #endregion
    }

}




