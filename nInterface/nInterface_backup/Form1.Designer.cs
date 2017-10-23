namespace nInterface
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnERPChk = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnMESChk = new System.Windows.Forms.Button();
            this.txtERPDB = new System.Windows.Forms.TextBox();
            this.txtMESDB = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnEnd = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbMailCheck = new System.Windows.Forms.ComboBox();
            this.cbSMSCheck = new System.Windows.Forms.ComboBox();
            this.cbTimer = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.lbldbLog = new System.Windows.Forms.Label();
            this.lbldbQuery = new System.Windows.Forms.Label();
            this.lblPath = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.listERP = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.listMES = new System.Windows.Forms.ListView();
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lblTime = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.mainTimer = new System.Windows.Forms.Timer(this.components);
            this.label9 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnERPChk
            // 
            this.btnERPChk.Location = new System.Drawing.Point(131, 32);
            this.btnERPChk.Name = "btnERPChk";
            this.btnERPChk.Size = new System.Drawing.Size(85, 25);
            this.btnERPChk.TabIndex = 0;
            this.btnERPChk.Text = "ERP Check";
            this.btnERPChk.UseVisualStyleBackColor = true;
            this.btnERPChk.Click += new System.EventHandler(this.ERPConnection_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.label1.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(418, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "  DB 접속정보                                                                 ";
            // 
            // btnMESChk
            // 
            this.btnMESChk.Location = new System.Drawing.Point(338, 33);
            this.btnMESChk.Name = "btnMESChk";
            this.btnMESChk.Size = new System.Drawing.Size(85, 25);
            this.btnMESChk.TabIndex = 3;
            this.btnMESChk.Text = "MES Check";
            this.btnMESChk.UseVisualStyleBackColor = true;
            this.btnMESChk.Click += new System.EventHandler(this.MESConnection_Click);
            // 
            // txtERPDB
            // 
            this.txtERPDB.Location = new System.Drawing.Point(20, 34);
            this.txtERPDB.Name = "txtERPDB";
            this.txtERPDB.ReadOnly = true;
            this.txtERPDB.Size = new System.Drawing.Size(100, 21);
            this.txtERPDB.TabIndex = 4;
            this.txtERPDB.Text = "Disconnection";
            // 
            // txtMESDB
            // 
            this.txtMESDB.Location = new System.Drawing.Point(232, 35);
            this.txtMESDB.Name = "txtMESDB";
            this.txtMESDB.ReadOnly = true;
            this.txtMESDB.Size = new System.Drawing.Size(100, 21);
            this.txtMESDB.TabIndex = 5;
            this.txtMESDB.Text = "Disconnection";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.label2.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label2.Location = new System.Drawing.Point(12, 87);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(420, 17);
            this.label2.TabIndex = 6;
            this.label2.Text = "  기준정보                                                                      ";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(257, 666);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(85, 25);
            this.btnStart.TabIndex = 7;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.Start_Click);
            // 
            // btnEnd
            // 
            this.btnEnd.Location = new System.Drawing.Point(347, 666);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(85, 25);
            this.btnEnd.TabIndex = 8;
            this.btnEnd.Text = "Stop";
            this.btnEnd.UseVisualStyleBackColor = true;
            this.btnEnd.Click += new System.EventHandler(this.Stop_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 113);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "오류시 메일발송";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 137);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 12);
            this.label4.TabIndex = 10;
            this.label4.Text = "오류시 SMS발송";
            // 
            // cbMailCheck
            // 
            this.cbMailCheck.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbMailCheck.FormattingEnabled = true;
            this.cbMailCheck.ImeMode = System.Windows.Forms.ImeMode.Katakana;
            this.cbMailCheck.Items.AddRange(new object[] {
            "Y",
            "N"});
            this.cbMailCheck.Location = new System.Drawing.Point(129, 110);
            this.cbMailCheck.Name = "cbMailCheck";
            this.cbMailCheck.Size = new System.Drawing.Size(64, 20);
            this.cbMailCheck.TabIndex = 11;
            // 
            // cbSMSCheck
            // 
            this.cbSMSCheck.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSMSCheck.FormattingEnabled = true;
            this.cbSMSCheck.Items.AddRange(new object[] {
            "Y",
            "N"});
            this.cbSMSCheck.Location = new System.Drawing.Point(129, 134);
            this.cbSMSCheck.Name = "cbSMSCheck";
            this.cbSMSCheck.Size = new System.Drawing.Size(64, 20);
            this.cbSMSCheck.TabIndex = 12;
            // 
            // cbTimer
            // 
            this.cbTimer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbTimer.FormattingEnabled = true;
            this.cbTimer.Items.AddRange(new object[] {
            "5",
            "10",
            "15",
            "20",
            "30",
            "40"});
            this.cbTimer.Location = new System.Drawing.Point(360, 113);
            this.cbTimer.Name = "cbTimer";
            this.cbTimer.Size = new System.Drawing.Size(64, 20);
            this.cbTimer.TabIndex = 14;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(231, 116);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(127, 12);
            this.label5.TabIndex = 13;
            this.label5.Text = "프로그램 실행간격(분)";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.label6.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label6.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label6.Location = new System.Drawing.Point(12, 179);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(419, 17);
            this.label6.TabIndex = 15;
            this.label6.Text = "  인터페이스 정보                                                             ";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(13, 211);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(78, 12);
            this.label7.TabIndex = 16;
            this.label7.Text = "ERP -> MES";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label11.Location = new System.Drawing.Point(18, 644);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(337, 12);
            this.label11.TabIndex = 30;
            this.label11.Text = "[기타] 경로변경은 \"App.config\" 파일을 참고 바랍니다.";
            // 
            // lbldbLog
            // 
            this.lbldbLog.AutoSize = true;
            this.lbldbLog.Location = new System.Drawing.Point(162, 590);
            this.lbldbLog.Name = "lbldbLog";
            this.lbldbLog.Size = new System.Drawing.Size(97, 12);
            this.lbldbLog.TabIndex = 29;
            this.lbldbLog.Text = "DB정보 파일경로";
            // 
            // lbldbQuery
            // 
            this.lbldbQuery.AutoSize = true;
            this.lbldbQuery.Location = new System.Drawing.Point(162, 570);
            this.lbldbQuery.Name = "lbldbQuery";
            this.lbldbQuery.Size = new System.Drawing.Size(97, 12);
            this.lbldbQuery.TabIndex = 28;
            this.lbldbQuery.Text = "DB정보 파일경로";
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(162, 549);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(97, 12);
            this.lblPath.TabIndex = 27;
            this.lblPath.Text = "DB정보 파일경로";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(18, 590);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(105, 12);
            this.label15.TabIndex = 26;
            this.label15.Text = "각종로그 파일경로";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(18, 570);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(97, 12);
            this.label16.TabIndex = 25;
            this.label16.Text = "DB쿼리 파일경로";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(18, 549);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(97, 12);
            this.label17.TabIndex = 24;
            this.label17.Text = "DB정보 파일경로";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.label18.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label18.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label18.Location = new System.Drawing.Point(12, 521);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(420, 17);
            this.label18.TabIndex = 23;
            this.label18.Text = "  기타정보                                                                      ";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(227, 211);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(78, 12);
            this.label8.TabIndex = 32;
            this.label8.Text = "MES -> ERP";
            // 
            // listERP
            // 
            this.listERP.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this.listERP.Font = new System.Drawing.Font("맑은 고딕", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.listERP.Location = new System.Drawing.Point(13, 230);
            this.listERP.Name = "listERP";
            this.listERP.Size = new System.Drawing.Size(201, 256);
            this.listERP.TabIndex = 34;
            this.listERP.UseCompatibleStateImageBehavior = false;
            this.listERP.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "No";
            this.columnHeader1.Width = 30;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Table Name";
            this.columnHeader2.Width = 200;
            // 
            // listMES
            // 
            this.listMES.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader3,
            this.columnHeader4});
            this.listMES.Font = new System.Drawing.Font("맑은 고딕", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.listMES.Location = new System.Drawing.Point(229, 230);
            this.listMES.Name = "listMES";
            this.listMES.Size = new System.Drawing.Size(201, 256);
            this.listMES.TabIndex = 35;
            this.listMES.UseCompatibleStateImageBehavior = false;
            this.listMES.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "No";
            this.columnHeader3.Width = 30;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Table Name";
            this.columnHeader4.Width = 200;
            // 
            // lblTime
            // 
            this.lblTime.AutoSize = true;
            this.lblTime.Location = new System.Drawing.Point(311, 500);
            this.lblTime.Name = "lblTime";
            this.lblTime.Size = new System.Drawing.Size(0, 12);
            this.lblTime.TabIndex = 37;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(205, 500);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(109, 12);
            this.label10.TabIndex = 36;
            this.label10.Text = "마지막 실행시간 :  ";
            // 
            // mainTimer
            // 
            this.mainTimer.Tick += new System.EventHandler(this.mainTimer_Tick);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(162, 611);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(97, 12);
            this.label9.TabIndex = 39;
            this.label9.Text = "DB정보 파일경로";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(18, 611);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(105, 12);
            this.label12.TabIndex = 38;
            this.label12.Text = "기준정보 파일경로";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(441, 703);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.lblTime);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.listMES);
            this.Controls.Add(this.listERP);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.lbldbLog);
            this.Controls.Add(this.lbldbQuery);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.cbTimer);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cbSMSCheck);
            this.Controls.Add(this.cbMailCheck);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtMESDB);
            this.Controls.Add(this.txtERPDB);
            this.Controls.Add(this.btnMESChk);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnERPChk);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(200, 100);
            this.MaximumSize = new System.Drawing.Size(457, 741);
            this.MinimumSize = new System.Drawing.Size(457, 741);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "ERP <-> MES Interface Program (SEMI)";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnERPChk;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnMESChk;
        private System.Windows.Forms.TextBox txtERPDB;
        private System.Windows.Forms.TextBox txtMESDB;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnEnd;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbMailCheck;
        private System.Windows.Forms.ComboBox cbSMSCheck;
        private System.Windows.Forms.ComboBox cbTimer;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label lbldbLog;
        private System.Windows.Forms.Label lbldbQuery;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ListView listERP;
        private System.Windows.Forms.ListView listMES;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.Label lblTime;
        private System.Windows.Forms.Label label10;
        //private System.Threading.Timer mainTimer;
        private System.Windows.Forms.Timer mainTimer;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label12;
    }
}

