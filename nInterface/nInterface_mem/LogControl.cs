using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Diagnostics;


namespace nInterface
{
    class LogControl
    {
        #region 클래스 멤버변수
        private string _sLogPath = string.Empty;
        #endregion

        #region 생성자
        public LogControl()
        {

        }
        public LogControl(string Path)
        {
            this._sLogPath = Path;
        }
        #endregion

        #region 멤버변수 대입
        public string LogPathString
        {
            get { return _sLogPath; }
            set { _sLogPath = value; }
        }
        #endregion

        #region IOFileWrite() 로그 저장
        /// <summary>
        /// 파라메터로 전달된 텍스트를 로그파일에 저장합니다.
        /// </summary>
        /// <param name="Gubun">구분입니다. (MES인지 ERP인지, 오류가 발생되었는지 등..)</param>
        /// <param name="ex">ini을 파싱 한 정보입니다. (메일주소, 휴대폰그룹 등)</param>
        /// <param name="querySource">조회하여 IF할 데이터를 만든 SELECT 쿼리 입니다. </param>
        /// <param name="queryTarget">IF할 데이터 전달하기 위해 실행 할 트랜젝션 쿼리입니다.</param>
        public void IOFileWrite(string Gubun, Exception ex, string querySource, string queryTarget)
        {
            //string logPath = string.Format("{0}\\{1}_Log.txt", _sLogPath, DateTime.Today.ToShortDateString());                               //저장될 로그 이름(로그가 생성된 날짜)
            //string errorLogPath = string.Format("{0}\\ERR\\{1}_ErrorLog.txt", _sLogPath, DateTime.Today.ToShortDateString());                //저장될 로그 이름(로그가 생성된 날짜)

            //declare folder of New Log
            string folderPart = string.Format("{0}\\ALL\\{1}", _sLogPath, DateTime.Now.ToString("yyyy-MM-dd"));
            string folderErrorPart = string.Format("{0}\\ERR\\{1}", _sLogPath, DateTime.Now.ToString("yyyy-MM-dd"));

            //declare route of Log data
            string logPath = string.Format("{0}\\ALL\\{1}\\{2}_Log.txt", _sLogPath, DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd_HH"));                               //저장될 로그 이름(로그가 생성된 날짜)
            string errorLogPath = string.Format("{0}\\ERR\\{1}\\{2}_ErrorLog.txt", _sLogPath, DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("yyyy-MM-dd_HH"));                //저장될 로그 이름(로그가 생성된 날짜)
            
            string errorFlag = string.Empty;
            StringBuilder logMsg = new StringBuilder();

            //logMsg.AppendLine(string.Format("[{0}]", DateTime.Now.ToString("yyyy/MM/dd/ HH:mm:ss:ffffff")));
            logMsg.AppendLine(string.Format("[{0}]", System.DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss':'ffffff")));
            
            //애러를 표시하는 괄호가 있으면
            if(Gubun.IndexOf("[") == -1)
            {
                errorFlag = "N"; //애러가 있을 경우는 별도의 파일에 애러로그만 따로 찍는다.

                logMsg.AppendLine(string.Format("   STATE  = {0}", ""));
                logMsg.AppendLine(string.Format("   TYPE   = {0}", Gubun));
            }
            else
            {
                errorFlag = "Y";

                logMsg.AppendLine(string.Format("   STATE = {0}", Gubun.Substring(Gubun.IndexOf("[")+1, 5) ));
                logMsg.AppendLine(string.Format("   TYPE = {0}", Gubun.Substring(0, Gubun.IndexOf("[")) ));
            }

            logMsg.AppendLine(string.Format("   MSG    = {0}", (ex == null) ? "" : ex.Message.ToString().Replace("\n", "")));
            logMsg.AppendLine(string.Format("   SOURCE_QUERY = {0}", querySource));
            logMsg.AppendLine(string.Format("   TARGET_QUERY = {0}", queryTarget));


            //create log folder
            if (!Directory.Exists(folderPart))         //Log를 기록할 폴더의 유무를 확인하여 생성
                Directory.CreateDirectory(folderPart);

            if (!Directory.Exists(folderErrorPart))         //Log를 기록할 폴더의 유무를 확인하여 생성
                Directory.CreateDirectory(folderErrorPart);


            //write log date
            //Log내용을 txt파일에 기록 (모든 로그파일)
            File.AppendAllText(logPath, logMsg.ToString());   

            //Log내용을 txt파일에 기록 (애러 로그파일)
            //기록 전 stackTrace까지 셋팅 후 기록
            if(errorFlag == "Y")
            {
                logMsg.AppendLine(string.Format("   STACKTRACE = {0}", ex.StackTrace.ToString())); //애러로그에 stack위치 출력 (오류 라인 등등)
                File.AppendAllText(errorLogPath, logMsg.ToString());
            }
        }
        #endregion

    }
}
