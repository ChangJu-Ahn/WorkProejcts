using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScrapWeightingEqp
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
        }

        DataTable _dtPlat = new DataTable();
        DataTable _dtEqp = new DataTable();
        string _sPlant = "";
        string _sEqpId = "";
        private void Setting_Load(object sender, EventArgs e)
        {
            InitiSet();
        }

        public DataTable PlantDt
        {
            set
            {
                _dtPlat = value;
            }
        }

        public DataTable EqpDt
        {
            set
            {
                _dtEqp = value;
            }
        }

        public string PlantCd
        {
            get
            {
                return _sPlant;
            }
            set
            {
                _sPlant = value;
            }
        }

        public string EqpId
        {
            get
            {
                return _sEqpId;
            }
            set
            {
                _sEqpId = value;
            }
        }


        private void InitiSet()
        {
            cmbPlat.DataSource = _dtPlat;
            cmbEQPID.DataSource = _dtEqp;

            cmbEQPID.SelectedValue = _sEqpId;
            //txtEqpId.Text = _sEqpId;

            cmbPlat.SelectedValue = _sPlant;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _sEqpId = cmbEQPID.SelectedValue.ToString();
            _sPlant = cmbPlat.SelectedValue.ToString();

            
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

    }
}
