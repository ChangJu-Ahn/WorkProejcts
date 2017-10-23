using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;

namespace nInterface
{
    class DBConn
    {

        #region 클래스 멤버변수
        private string _sConnInfo = string.Empty;
        private string _sConnERPInfo = string.Empty;
        private string _sConnMESInfo = string.Empty;

        OleDbConnection _sqlDBConn = null;
        OleDbConnection _sqlERPConn = null;
        OleDbConnection _sqlMESConn = null;
        #endregion

        #region 생성자
        public DBConn()
        {

        }
        public DBConn(string ConnInfo)
        {
            this._sConnInfo = ConnInfo;
            InitDBConn();
        }
        #endregion

        #region DB Connection
        /// <summary>
        /// 데이터베이스에 접속합니다.
        /// </summary>
        public void InitDBConn()
        {
            _sqlDBConn = new OleDbConnection(_sConnInfo);
            _sqlDBConn.Open();
        }
        #endregion

        #region 멤버변수 대입
        /// <summary>
        /// 멤버변수를 대입하거나 불러옵니다.
        /// </summary>
        public string ConnectionString
        {
            get { return _sConnInfo; }
            set { _sConnInfo = value; }
        }
        public string ConnectionERPString
        {
            get { return _sConnERPInfo; }
            set { _sConnERPInfo = value; }
        }
        public string ConnectionMESString
        {
            get { return _sConnMESInfo; }
            set { _sConnMESInfo = value; }
        }
        #endregion

        #region DB Close
        /// <summary>
        /// 데이터베이스 연결을 종료합니다.
        /// </summary>
        public void CloseDBConn()
        {
            if (_sqlDBConn != null && _sqlDBConn.State != ConnectionState.Closed)
            {
                _sqlDBConn.Close();
                _sqlDBConn = null;
            }
        }

        /// <summary>
        /// 데이터베이스 연결을 종료합니다.
        /// </summary>
        public void CloseERPDBConn()
        {
            if (_sqlERPConn != null && _sqlERPConn.State != ConnectionState.Closed)
            {
                _sqlERPConn.Close();
                _sqlERPConn = null;
            }
        }

        /// <summary>
        /// 데이터베이스 연결을 종료합니다.
        /// </summary>
        public void CloseMESDBConn()
        {
            if (_sqlMESConn != null && _sqlMESConn.State != ConnectionState.Closed)
            {
                _sqlMESConn.Close();
                _sqlMESConn = null;
            }
        }
        #endregion

        #region DB쿼리 실행(SELECT 실행)
        /// <summary>
        /// 트랜잭션이 없는 단순 조회쿼리를 실행합니다.
        /// </summary>
        /// <param name="query">실행되어야 할 문장(쿼리)</param>
        /// <param name="type">ERP쿼리인지 MES쿼리인지 판단하는 구분자 (객체를 open할 때 사용)</param>
        public DataTable ExecuteQuery(string query, string type)
        {
            //Connection Check
            DBConnectCheck(type.ToUpper());

            DataTable dt = new DataTable();
            OleDbCommand oleCmd;
            OleDbDataAdapter oleAdpt;

            //구분자를 확인하여 ERP객체 또는 MES객체를 선택하여 사용
            if (type.ToUpper() == "MES")
                oleCmd = new OleDbCommand(query, _sqlMESConn);
            else
                oleCmd = new OleDbCommand(query, _sqlERPConn);

            oleAdpt = new OleDbDataAdapter(oleCmd);
            oleCmd.CommandTimeout = 15; //테이블 Lock일 경우 무한정 대기상태에 빠지게되니 timeout을 걸어서 15초가 지날 경우 애러를 출력하고 다음프로세스를 진행
            oleAdpt.Fill(dt);

            return dt;
        }
        #endregion

        #region DB상태 및 연결하는 메소드입니다.
        /// <summary>
        /// ERP 객체(DB)를 연결 및 Open 합니다.
        /// </summary>
        public void InitERP_DBConn()
        {
            _sqlERPConn = new OleDbConnection(_sConnERPInfo);
            _sqlERPConn.Open();
        }

        /// <summary>
        /// MES 객체(DB)를 연결 및 Open 합니다.
        /// </summary>
        public void InitMES_DBConn()
        {
            _sqlMESConn = new OleDbConnection(_sConnMESInfo);
            _sqlMESConn.Open();
        }

        /// <summary>
        /// 데이터베이스의 연결상태를 확인합니다. (연결이 안 되어 있다면 연결합니다)
        /// </summary>
        /// <param name="type">어떤 객체를 연결할 것인지에 대한 구분자 입니다 (ERP, MES)<param>
        private void DBConnectCheck(string type)
        {
            switch (type)
            {
                case "MES":
                    if (_sqlMESConn == null || _sqlMESConn.State != ConnectionState.Open)
                        InitMES_DBConn();

                    break;

                case "ERP":
                    if (_sqlERPConn == null || _sqlERPConn.State != ConnectionState.Open)
                        InitERP_DBConn();

                    break;

                default:
                    if (_sqlERPConn == null || _sqlERPConn.State != ConnectionState.Open)
                        InitERP_DBConn();

                    if (_sqlMESConn == null || _sqlMESConn.State != ConnectionState.Open)
                        InitMES_DBConn();

                    break;
            }
        }
        #endregion

        #region DB쿼리 실행(트랜잭션 쿼리 실행 / insert, delete, update)
        /// <summary>
        /// 트랜젝션이 있는 DML쿼리를 실행합니다. (트랜잭션이 일어나는 insert, delete, updae)
        /// </summary>
        /// <param name="insertQuery">실행되어야 할 insert 문장<param>
        /// <param name="updateQuery">실행되어야 할 update 문장</param>
        /// <param name="type">ERP쿼리인지 MES쿼리인지 판단하는 구분자 (객체를 open할 때 사용)</param>
        public bool ExecuteTransactionNonQuery(string insertQuery, string updateQuery, string type)
        {
            //Connection Check
            DBConnectCheck(type.ToUpper());

            int nResultInsert = 0;
            int nResultUpdate = 0;

            //오라클과 MSSQL의 트랜젝션을 동시에 관리해야 하기 때문에 별도의 객체를 생성
            OleDbTransaction oleTranERP = _sqlERPConn.BeginTransaction();
            OleDbTransaction oleTranMES = _sqlMESConn.BeginTransaction();
            OleDbCommand oleCmdERP;
            OleDbCommand oleCmdMES;

            //구분자별 쿼리 셋팅
            switch (type.ToUpper())
            {
                case "MES":
                    //MES에서 시작되는 경우 ERP에 INSERT를 해야 하니 ERP객체에는 INSERT 쿼리를, MES에는 INSERT 후 업데이트 해야하니 UPDATE 쿼리를 셋팅
                    oleCmdERP = new OleDbCommand(insertQuery, _sqlERPConn);
                    oleCmdMES = new OleDbCommand(updateQuery, _sqlMESConn);

                    break;

                case "ERP":
                    //ERP에서 시작되는 경우 MES에 INSERT를 해야 하니 MES객체에는 INSERT 쿼리를, ERP에는 INSERT 후 업데이트 해야하니 UPDATE 쿼리를 셋팅
                    oleCmdMES = new OleDbCommand(insertQuery, _sqlMESConn);
                    oleCmdERP = new OleDbCommand(updateQuery, _sqlERPConn);

                    break;

                default:
                    return false;
            }

            //각 객체 별 트랜잭션 설정
            oleCmdERP.Transaction = oleTranERP;
            oleCmdMES.Transaction = oleTranMES;
            
            //각 객체 별 타임아웃 설정
            oleCmdERP.CommandTimeout = 8;
            oleCmdMES.CommandTimeout = 8;

            
            try
            {
                //트랜잭션 쿼리 실행 (INSERT 갯수는 트리거때문에 1개 이상이 나올 수 있지만 업데이트는 
                if (type.ToUpper() == "MES")
                {
                    nResultInsert = oleCmdERP.ExecuteNonQuery();
                    nResultUpdate = oleCmdMES.ExecuteNonQuery();
                }
                else
                {
                    nResultUpdate = oleCmdERP.ExecuteNonQuery();
                    nResultInsert = oleCmdMES.ExecuteNonQuery();
                }

                //한개 이상 업데이트가 되거나 한개 이상 insert가 될 경우는 롤백처리
                //UPDATE는 한개씩만 되어야 하지만 테이블에 트리거가 걸려있을 경우 한 개의 로우에만 UPDATE가 되었어도 테이블이 많으므로 1개 이상을 리턴할 수 있음
                //if (nResultInsert == 0 && nResultUpdate == 0)
                if (nResultInsert >= 1 && nResultUpdate >= 1)
                {
                    oleTranERP.Commit();
                    oleTranMES.Commit();

                    return true;
                }
                else
                {
                    oleTranERP.Rollback();
                    oleTranMES.Rollback();

                    return false;
                }
            }
            catch (Exception ex)
            {
                oleTranERP.Rollback();
                oleTranMES.Rollback();

                //애러를 상위단으로 던짐
                throw new Exception(ex.Message);
                
            }
        }


        /// <summary>
        /// DML쿼리를 실행합니다. (트랜잭션이 일어나는 insert, delete, updae)
        /// </summary>
        /// <param name="query">실행되어야 할 문장(쿼리)</param>
        /// <param name="type">ERP쿼리인지 MES쿼리인지 판단하는 구분자 (객체를 open할 때 사용)</param>
        public int ExecuteNonQuery(string query, string type)
        {
            //Connection Check
            DBConnectCheck(type.ToUpper());

            int nResult = 0;
            OleDbTransaction oleTran = _sqlDBConn.BeginTransaction(); //auto commit 해제 (begin Transaction 설정)
            OleDbCommand oleCmd = new OleDbCommand();

            try
            {
                //연결 및 쿼리 설정 or 트랜잭션을 설정하여 오류가 발생 될 경우 rollback 처리 실시
                oleCmd.CommandText = query;

                //구분자를 확인하여 ERP객체 또는 MES객체를 선택하여 사용
                if (type.ToUpper() == "MES")
                    oleCmd.Connection = _sqlMESConn;
                else
                    oleCmd.Connection = _sqlERPConn;

                oleCmd.Transaction = oleTran;
                oleCmd.CommandTimeout = 5; //테이블 Lock일 경우 무한정 대기상태에 빠지게되니 timeout을 설정(Row별로 진행해야 하기 때문에 15초에서 5초로 변경)

                //트랜잭션 쿼리 실행
                nResult = oleCmd.ExecuteNonQuery();

                //트랜잭션 후 이상없으면 Commit
                oleTran.Commit();

                return nResult;
            }
            catch (Exception ex)
            {
                oleTran.Rollback();

                //애러를 상위단으로 던짐
                throw new Exception(ex.Message);
            }
        }
        #endregion

    }
}
