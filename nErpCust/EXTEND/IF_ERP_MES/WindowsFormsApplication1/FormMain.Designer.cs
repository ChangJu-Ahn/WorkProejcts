namespace T_IF_RCV_PROD_ORD_KO441
{
    partial class FormMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.btnTimer = new System.Windows.Forms.Button();
            this.btnMSSql = new System.Windows.Forms.Button();
            this.btnOracle = new System.Windows.Forms.Button();
            this.txtMSSql = new System.Windows.Forms.TextBox();
            this.txtOracle = new System.Windows.Forms.TextBox();
            this.mainTimer = new System.Windows.Forms.Timer(this.components);
            this.btnManual = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // btnTimer
            // 
            this.btnTimer.BackColor = System.Drawing.Color.Red;
            this.btnTimer.Location = new System.Drawing.Point(4, 26);
            this.btnTimer.Name = "btnTimer";
            this.btnTimer.Size = new System.Drawing.Size(519, 61);
            this.btnTimer.TabIndex = 1;
            this.btnTimer.Text = "Timer Start";
            this.btnTimer.UseVisualStyleBackColor = false;
            this.btnTimer.Click += new System.EventHandler(this.btnTimer_Click);
            // 
            // btnMSSql
            // 
            this.btnMSSql.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnMSSql.Location = new System.Drawing.Point(4, 93);
            this.btnMSSql.Name = "btnMSSql";
            this.btnMSSql.Size = new System.Drawing.Size(253, 61);
            this.btnMSSql.TabIndex = 3;
            this.btnMSSql.Text = "MSSQL (ERP)";
            this.btnMSSql.UseVisualStyleBackColor = false;
            this.btnMSSql.Click += new System.EventHandler(this.btnMSSql_Click);
            // 
            // btnOracle
            // 
            this.btnOracle.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnOracle.Location = new System.Drawing.Point(270, 93);
            this.btnOracle.Name = "btnOracle";
            this.btnOracle.Size = new System.Drawing.Size(253, 61);
            this.btnOracle.TabIndex = 4;
            this.btnOracle.Text = "Oracle (MES)";
            this.btnOracle.UseVisualStyleBackColor = false;
            this.btnOracle.Click += new System.EventHandler(this.btnOracle_Click);
            // 
            // txtMSSql
            // 
            this.txtMSSql.Location = new System.Drawing.Point(4, 161);
            this.txtMSSql.Name = "txtMSSql";
            this.txtMSSql.Size = new System.Drawing.Size(253, 21);
            this.txtMSSql.TabIndex = 5;
            // 
            // txtOracle
            // 
            this.txtOracle.Location = new System.Drawing.Point(270, 160);
            this.txtOracle.Name = "txtOracle";
            this.txtOracle.Size = new System.Drawing.Size(253, 21);
            this.txtOracle.TabIndex = 6;
            // 
            // mainTimer
            // 
            //this.mainTimer.Interval = 30000;
            this.mainTimer.Tick += new System.EventHandler(this.mainTimer_Tick);
            // 
            // btnManual
            // 
            this.btnManual.Location = new System.Drawing.Point(4, 189);
            this.btnManual.Name = "btnManual";
            this.btnManual.Size = new System.Drawing.Size(519, 39);
            this.btnManual.TabIndex = 7;
            this.btnManual.Text = "메뉴얼 실행";
            this.btnManual.UseVisualStyleBackColor = true;
            this.btnManual.Click += new System.EventHandler(this.btnManual_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("굴림", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.richTextBox1.ForeColor = System.Drawing.Color.Red;
            this.richTextBox1.Location = new System.Drawing.Point(4, 234);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(519, 126);
            this.richTextBox1.TabIndex = 8;
            this.richTextBox1.Text = resources.GetString("richTextBox1.Text");
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(529, 372);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.btnManual);
            this.Controls.Add(this.txtOracle);
            this.Controls.Add(this.txtMSSql);
            this.Controls.Add(this.btnOracle);
            this.Controls.Add(this.btnMSSql);
            this.Controls.Add(this.btnTimer);
            this.Name = "FormMain";
            this.Text = "ERP I/F (작업지시,라우팅,자재입출고)";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnTimer;
        private System.Windows.Forms.Button btnMSSql;
        private System.Windows.Forms.Button btnOracle;
        private System.Windows.Forms.TextBox txtMSSql;
        private System.Windows.Forms.TextBox txtOracle;
        private System.Windows.Forms.Timer mainTimer;
        private System.Windows.Forms.Button btnManual;
        private System.Windows.Forms.RichTextBox richTextBox1;
    }
}

