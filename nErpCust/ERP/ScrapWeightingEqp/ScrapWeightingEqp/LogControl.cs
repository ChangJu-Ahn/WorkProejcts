using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace ScrapWeightingEqp
{
    class LogControl
    {
        #region 클래스 멤버변수
        private string _sLogPath = string.Empty;
        private string _sDbInfo = string.Empty;
        #endregion

        #region 생성자
        public LogControl()
        {

        }
        public LogControl(string Path, string DB)
        {
            this._sLogPath = Path;
            this._sDbInfo = DB;
        }
        #endregion

        #region 멤버변수 대입
        public string LogPathString
        {
            get { return _sLogPath; }
            set { _sLogPath = value; }
        }
        #endregion

        public string GetDateTime()
        {
            DateTime NowDate = DateTime.Now;
            return NowDate.ToString("yyyy-MM-dd HH:mm:ss") + ":" + NowDate.Millisecond.ToString("000");
        }

        private string GetLogStatus(int sts)
        {
            string sResult = "";

            if(sts == 1)
            {
             sResult = "INFO".PadRight(8, ' ');
            }
            else if(sts == 2)
            {
             sResult = "SUCCESS".PadRight(8, ' ');
            }
            else if(sts == 3)
            {
             sResult = "ERROR".PadRight(8, ' ');
            }

            return sResult;
        }


        #region 로그 저장
        public void IOFileWrite(string msg, int logsts)
        {
            //string dicPath = string.Format("{0}\\Log", _sLogPath);                                                          //저장될 로그 경로
            //string logPath = string.Format("{0}\\{1}_Log.txt", dicPath, DateTime.Today.ToShortDateString());                //저장될 로그 이름(로그가 생성된 날짜)
            //string logMsg = string.Format("[{0}] {1} {2}", DateTime.Now.ToString("u"), msg, Environment.NewLine); //저장될 로그형식 : [2016-05-11 11:30:005] 로그데이터 출력 테스트

            //if (!Directory.Exists(dicPath))         //Log를 기록할 폴더의 유무를 확인하여 생성
            //    Directory.CreateDirectory(dicPath);

            //File.AppendAllText(logPath, logMsg);        //Log내용을 txt파일에 기록



            string sDirPath = _sLogPath + "\\Logs";
            string sLogPath = sDirPath + "\\Log" + DateTime.Today.ToString("yyyyMMdd") + ".log";
            
            string sTemp;

            DirectoryInfo di = new DirectoryInfo(sDirPath);
            FileInfo fi = new FileInfo(sLogPath);

            try
            {
                if (di.Exists != true)
                {
                    Directory.CreateDirectory(sDirPath);
                }
                if (fi.Exists != true)
                {
                    using (StreamWriter sw = new StreamWriter(sLogPath))
                    {
                        sTemp = string.Format("[{0}][{1}][{2}]:{3}", GetDateTime(), _sDbInfo.PadRight(14, ' '), GetLogStatus(logsts), msg);
                        sw.WriteLine(sTemp);
                        sw.Close();
                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(sLogPath))
                    {
                        sTemp = string.Format("[{0}][{1}][{2}]:{3}", GetDateTime(), _sDbInfo.PadRight(14, ' '), GetLogStatus(logsts), msg);
                        sw.WriteLine(sTemp);
                        sw.Close();
                    }
                }
            }
            catch (Exception ex)
            {

            }


        }
        #endregion

    }
}
