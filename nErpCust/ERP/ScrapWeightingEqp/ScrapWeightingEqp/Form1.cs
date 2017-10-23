using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Diagnostics;

namespace ScrapWeightingEqp
{
    public partial class Form1 : Form
    {

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect, // x-coordinate of upper-left corner
            int nTopRect, // y-coordinate of upper-left corner
            int nRightRect, // x-coordinate of lower-right corner
            int nBottomRect, // y-coordinate of lower-right corner
            int nWidthEllipse, // height of ellipse
            int nHeightEllipse // width of ellipse
         );


        public Form1()
        {
            InitializeComponent();
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 10, 10));
        }

        public enum LogStatus : int
        {
            Info = 1
            ,
            Success = 2
           , Error = 3

        }

        private static string _dbInfo = System.Configuration.ConfigurationSettings.AppSettings["Nepes_DB"];
        private static string _saveLogPath = System.Configuration.ConfigurationSettings.AppSettings["Log_Path"];
        private static string _sTextPath = System.Configuration.ConfigurationSettings.AppSettings["Text_Path"];
        private static string _sPlant = System.Configuration.ConfigurationSettings.AppSettings["Plant"];
        private static string _sEqpId = System.Configuration.ConfigurationSettings.AppSettings["EqpID"];


        private Point mousePoint;
     
        static DBConn sqlConn = null;
        static LogControl logCtrl = null;
        private bool chk = false;

        DateTime dtLast = new DateTime();
        System.Threading.Timer _timer;
        Thread workThread = null;
        delegate void TimerEventFiredDelegate();
        private void Callback(object status)
        {
            // UI 에서 사용할 경우는 Cross-Thread 문제가 발생하므로 Invoke 또는 BeginInvoke 를 사용해서 마샬링을 통한 호출을 처리하여야 한다.
            BeginInvoke(new TimerEventFiredDelegate(Work));
        }
        private void CallbackErr(object status)
        {
            // UI 에서 사용할 경우는 Cross-Thread 문제가 발생하므로 Invoke 또는 BeginInvoke 를 사용해서 마샬링을 통한 호출을 처리하여야 한다.
            BeginInvoke(new TimerEventFiredDelegate(WorkErr));
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            Process[] processList = Process.GetProcessesByName("ScrapWeightingEqp");   //this is Program process name

            if (processList.Length > 1)
            {
                MessageBox.Show("이미 프로세스가 실행 중 입니다.");
                Application.ExitThread();
                Environment.Exit(0);
            }
            
            logCtrl = new LogControl(_saveLogPath, "Nepes_DB");

            //창 줄이기
            FormSize(0);

            chk = true;

            workThread = new Thread(Processing);

            workThread.Start();
            
        }

        private Boolean chkStart()
        {
            Boolean result = false;

           

            return result;
        }

        private void FormSize(int UpDn)
        {
            if(UpDn == 0)
            {
                splitContainer1.Panel2Collapsed = true;
                this.Size = new Size(this.Size.Width, 122);
                btnView.Text = "▽";
            }
            else
            {
                splitContainer1.Panel2Collapsed = false;
                this.Size = new Size(this.Size.Width, 345);
                btnView.Text = "△";
            }
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(5, 5, Width, Height, 5, 5));
        }

        
        private void Processing()
        {



            while (chk)
            {

                Thread.Sleep(1500);

                if (_sEqpId == "Test1")
                {
                    if (_timer != null)
                    {
                        _timer.Dispose();
                        _timer = null;
                    }

                    _timer = new System.Threading.Timer(CallbackErr);
                    _timer.Change(1000, 1000);    // dueTime 은 Timer 가 시작되기 전 대기 시간 (ms)
                    SetControl(lblDisplay, "설정에서 저울 ID를 설정해 주세요.", Color.Red);

                }
                else
                {

                    FileInfo fi = new FileInfo(_sTextPath + "Com_3.txt");

                    if (fi.Exists)
                    {

                        DateTime dt = fi.LastWriteTime;

                        if (dtLast != dt)
                        {

                            if (_timer != null)
                            {
                                _timer.Dispose();
                                _timer = null;
                            }

                            FileStream fs = new FileStream(_sTextPath + "Com_3.txt", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                            StreamReader sr = new StreamReader(fs);

                            string str = "";

                            str = sr.ReadToEnd();

                            dtLast = dt;

                            string[] arrstr = str.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                            SetControl(txtFile, str);
                            if (arrstr.Length > 0)
                            {
                                string sValue = arrstr[arrstr.Length - 1];
                                if (sValue.Length > 0)
                                {
                                    CheckStrValure(sValue);
                                    SetControl(textBox1, sValue);
                                }
                                else if (arrstr.Length >= 2)
                                {
                                    sValue = arrstr[arrstr.Length - 2];
                                    CheckStrValure(sValue);
                                    SetControl(textBox1, sValue);
                                }

                                //

                            }
                        }

                    }
                    else
                    {
                        if (_timer == null)
                        {
                            _timer = new System.Threading.Timer(CallbackErr);
                            _timer.Change(1000, 1000);    // dueTime 은 Timer 가 시작되기 전 대기 시간 (ms)
                            SetControl(lblDisplay, "저울 파일 미존재", Color.Red);
                        }
                        else
                        {
                            //_timer.Dispose();
                            //_timer = null;
                        }
                    }
                }
            }

        }

        private void Work()
        {
            if(lblDisplay.BackColor != Color.Red)
            {
                SetControl(lblDisplay, lblDisplay.Text, Color.White);
                _timer.Dispose();
                _timer = null;
            }
        }

        private void WorkErr()
        {
            if (lblDisplay.BackColor == Color.Red)
            {
                SetControl(lblDisplay, lblDisplay.Text, Color.White);
            }
            else
            {
                SetControl(lblDisplay, lblDisplay.Text, Color.Red);
            }
        }

        private Boolean CheckStrValure(string sValue)
        {
            Boolean result = false;
            string[] chkValure = sValue.Split(',');

            string sWeight1 = chkValure[1].Substring(0, 1);
            string sWeight2 = chkValure[1].Substring(1, chkValure[1].Length-1);

            double Weight = new double();

            if(chkValure.Length != 6)
            {
                _timer = new System.Threading.Timer(CallbackErr);
                _timer.Change(1000, 1000);    // dueTime 은 Timer 가 시작되기 전 대기 시간 (ms)
                SetControl(lblDisplay, "파일 형식 불일치", Color.Red);
                
            }
            else if (!double.TryParse(sWeight2, out Weight))
            {
                _timer = new System.Threading.Timer(CallbackErr);
                _timer.Change(1000, 1000);    // dueTime 은 Timer 가 시작되기 전 대기 시간 (ms)
                SetControl(lblDisplay, "무게정보 오류", Color.Red);
                
            }
            else
            {
                if (SetDbInsert(sValue))
                {
                    _timer = new System.Threading.Timer(Callback);
                    _timer.Change(10000, 1000);    // dueTime 은 Timer 가 시작되기 전 대기 시간 (ms)

                    result = true;

                    SetControl(lblDisplay, chkValure[1] + chkValure[2], Color.Green);

                }
                else
                {
                    _timer = new System.Threading.Timer(CallbackErr);
                    _timer.Change(1000, 1000);    // dueTime 은 Timer 가 시작되기 전 대기 시간 (ms)
                    SetControl(lblDisplay, "DB 저장 오류", Color.Red);
                }
            }


            return result;
        }

        delegate void DSetControl(Control ctl, string sText, Color clr);

        private void SetControl(Control ctl, string sText, Color clr)
        {
            if (this.InvokeRequired)
            {
                DSetControl d = new DSetControl(SetControl);
                this.Invoke(d, new object[] { ctl, sText, clr });

            }
            else
            {
                ctl.Text = sText;
                ctl.BackColor = clr;
            }

        }

        //delegate void DSetControl(Control ctl, string sText);

        private void SetControl(Control ctl, string sText)
        {
            if (this.InvokeRequired)
            {
                DSetControl d = new DSetControl(SetControl);
                this.Invoke(d, new object[] { ctl, sText, Color.White });

            }
            else
            {
                ctl.Text = sText;
            }

        }


        private Boolean SetDbInsert(string sValue)
        {
            StringBuilder sSQL = new StringBuilder();

            Boolean result = false;

            logCtrl.IOFileWrite(" DB Insert 시작 "+ sValue, (int)LogStatus.Info);

            string[] sDbValure = sValue.Split(',');

            if (sqlConn == null)
                sqlConn = new DBConn(_dbInfo);

            sSQL.AppendLine("INSERT INTO OUT_MAT_WEIGHT");
            sSQL.AppendLine("  (");
            sSQL.AppendLine("    TRANS_TIME ");
            sSQL.AppendLine("   , PLANT_CD ");
            sSQL.AppendLine("   , EQUIPMENT_ID ");
            sSQL.AppendLine("   , WEIGHT_VALUE ");
            sSQL.AppendLine("   , WEIGHT_UNIT ");
            sSQL.AppendLine("   , WEIGHT_TIME ");
            sSQL.AppendLine(" ) ");
            sSQL.AppendLine(" VALUES ");
            sSQL.AppendLine("  ( ");
            sSQL.AppendLine("   CONVERT(CHAR(8),GETDATE(),112)+REPLACE(CONVERT(CHAR(8),GETDATE(),108),':','')");
            sSQL.AppendLine("   , '" + _sPlant + "'");
            sSQL.AppendLine("   , '" + _sEqpId + "'");
            sSQL.AppendLine("   , " + sDbValure[1].Trim() + "");
            sSQL.AppendLine("   , '" + sDbValure[2].Trim() + "'");
            sSQL.AppendLine("   , CONVERT(DATETIME, '" + sDbValure[4] + " " + sDbValure[3] + "')");
            sSQL.AppendLine("  )");

            try
            {
                sqlConn.ExecuteQuery(sSQL.ToString());
                logCtrl.IOFileWrite(sSQL.ToString(), (int)LogStatus.Info);
                logCtrl.IOFileWrite(" [" + sValue + "] DB Insert 완료", (int)LogStatus.Success);
                result = true;
                
            }
            catch (Exception ex)
            {
                result = false;
                logCtrl.IOFileWrite(" [" + sValue + "]DB Insert ERROR" + ex.Message, (int)LogStatus.Error);
            }
            return result;
            
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(workThread != null)
            {
                workThread.Abort();
            }
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            if(splitContainer1.Panel2Collapsed)
            {
                FormSize(1);
            }
            else
            {
                FormSize(0);
            }
        }

        private void splitContainer1_Panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                Location = new Point(this.Left - (mousePoint.X - e.X),
                    this.Top - (mousePoint.Y - e.Y));
            }
        }

        private void splitContainer1_Panel1_MouseDown(object sender, MouseEventArgs e)
        {

            mousePoint = new Point(e.X, e.Y);
        }

        private void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            this.Opacity = Convert.ToDouble(trackBar1.Value*0.1);
        }

        private void chkTop_CheckedChanged(object sender, EventArgs e)
        {
            if(chkTop.Checked)
            {
                this.TopMost = true;
            }
            else
            {
                this.TopMost = false;
            }
        }

        private void lblDisplay_MouseDown(object sender, MouseEventArgs e)
        {
            mousePoint = new Point(e.X, e.Y);
        }

        private void lblDisplay_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                Location = new Point(this.Left - (mousePoint.X - e.X),
                    this.Top - (mousePoint.Y - e.Y));
            }
        }

        private void btnSetting_Click(object sender, EventArgs e)
        {
            Setting setfrm = new Setting();
            setfrm.PlantDt = GetPlantDt();
            setfrm.EqpDt = GetEQPDt();
            setfrm.PlantCd = _sPlant;
            setfrm.EqpId = _sEqpId;

            setfrm.StartPosition = FormStartPosition.CenterParent;
            setfrm.TopMost = this.TopMost;

            if(setfrm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _sEqpId = setfrm.EqpId;
                _sPlant = setfrm.PlantCd;

                //var appset =System.Configuration.

                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;

                settings["EqpID"].Value = setfrm.EqpId;
                settings["Plant"].Value = setfrm.PlantCd;

                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);

                _sPlant = System.Configuration.ConfigurationSettings.AppSettings["Plant"];
                _sEqpId = System.Configuration.ConfigurationSettings.AppSettings["EqpID"];

                logCtrl.IOFileWrite(" [" + _sEqpId + "] 저울 ID 변경 성공", (int)LogStatus.Success);
                logCtrl.IOFileWrite(" [" + _sPlant + "] 공장 정보 변경 성공", (int)LogStatus.Success);
            }

        }


        private DataTable GetPlantDt()
        {
            DataTable dt = new DataTable();
            StringBuilder sSQL = new StringBuilder();

            if (sqlConn == null)
                sqlConn = new DBConn(_dbInfo);

            sSQL.AppendLine("SELECT DISTINCT PLANT_CD, PLANT_DESC FROM DBO.SA_SYS_CODE");

            try
            {
                dt = sqlConn.ExecuteQuery(sSQL.ToString());
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(ex.Message, (int)LogStatus.Error);
            }

            return dt;
        }

        private DataTable GetEQPDt()
        {
            DataTable dt = new DataTable();
            StringBuilder sSQL = new StringBuilder();

            if (sqlConn == null)
                sqlConn = new DBConn(_dbInfo);

            sSQL.AppendLine("SELECT UD_MINOR_CD AS ITEM_CD , UD_MINOR_NM AS ITEM_NM FROM B_USER_DEFINED_MINOR WHERE UD_MAJOR_CD = 'SA003'");

            try
            {
                dt = sqlConn.ExecuteQuery(sSQL.ToString());
            }
            catch (Exception ex)
            {
                logCtrl.IOFileWrite(ex.Message, (int)LogStatus.Error);
            }

            return dt;
        }

  
    }
}
