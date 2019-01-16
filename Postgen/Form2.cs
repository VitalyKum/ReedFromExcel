using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;


namespace Postgen
{
    public partial class frmCorrect : Form
    {
        public Excel.Worksheet tSheet;
        public Excel.Worksheet sSheet;
        public int rCount = 0;
        private int rPos = 2;

        public frmCorrect()
        {
            InitializeComponent();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            cmdBack.Enabled = false;
            rPos = 2;
            if (rCount == 1)
                cmdForward.Enabled = false;

            FillCurrentRowData();
        }

        private string GetCurrentRowString(int pos, int max)
        {
            return string.Format("Запись {0} из {1}", (pos-1).ToString(), max.ToString());
        }

        private void FillCurrentRowData()
        {
            try
            {
                string oldID = string.IsNullOrEmpty(((Excel.Range)sSheet.Cells[rPos, 9]).Value) ? "" : string.Format("{0}", ((Excel.Range)sSheet.Cells[rPos, 9]).Value);
                string oldTel = string.IsNullOrEmpty(((Excel.Range)sSheet.Cells[rPos, 8]).Value) ? "" : string.Format("{0}", ((Excel.Range)sSheet.Cells[rPos, 8]).Value);
                string oldIndx = string.IsNullOrEmpty(((Excel.Range)sSheet.Cells[rPos, 7]).Value) ? "" : string.Format("{0}", ((Excel.Range)sSheet.Cells[rPos, 7]).Value);
                string oldAdr = string.IsNullOrEmpty(((Excel.Range)sSheet.Cells[rPos, 3]).Value) ? "" : string.Format("{0}", ((Excel.Range)sSheet.Cells[rPos, 3]).Value);
                string oldFio =  string.IsNullOrEmpty(((Excel.Range)sSheet.Cells[rPos, 6]).Value) ? "" : string.Format("{0}", ((Excel.Range)sSheet.Cells[rPos, 6]).Value);
                string mass = string.IsNullOrEmpty(((Excel.Range)tSheet.Cells[rPos, 3]).Value) ? "0.0" : string.Format("{0}", ((Excel.Range)tSheet.Cells[rPos, 3]).Value);
               

                txtoldID.Text = oldID;
                txtnewID.Text = ((Excel.Range)tSheet.Cells[rPos, 6]).Value;

                txtMass.Text = mass;
               
                txtoldIndx.Text = oldIndx;
                //txtnewIndx.Text = string.Format("{0}", ((Excel.Range)tSheet.Cells[rPos, 5]).Value);

                txtoldTel.Text = oldTel;
                txtnewTel.Text = string.Format("{0}", ((Excel.Range)tSheet.Cells[rPos, 7]).Value);

                txtoldFio.Text =oldFio;
                txtnewFio.Text = ((Excel.Range)tSheet.Cells[rPos, 2]).Value;

                txtoldAdr.Text = oldAdr;
                txtnewAdr.Text = ((Excel.Range)tSheet.Cells[rPos, 1]).Value;
            }
            catch (Exception err) when (err.Data != null)
            {
                MessageBox.Show(string.Format("Непредвиденная ошибка коррекции при чтении в строке {0}: {1}", rPos, err.Message));
            }

            lbCurrentRow.Text = GetCurrentRowString(rPos, rCount);
        }

        private void UpdateCurrentRowData()
        {
            try
            {               
                ((Excel.Range)tSheet.Cells[rPos, 6]).Value = txtnewID.Text;
                ((Excel.Range)tSheet.Cells[rPos, 3]).Value = txtMass.Text;                
                //((Excel.Range)tSheet.Cells[rPos, 5]).Value = txtnewIndx.Text;               
                ((Excel.Range)tSheet.Cells[rPos, 7]).Value = txtnewTel.Text;                
                ((Excel.Range)tSheet.Cells[rPos, 2]).Value = txtnewFio.Text;                
                ((Excel.Range)tSheet.Cells[rPos, 1]).Value = txtnewAdr.Text;
            }
            catch (Exception err) when (err.Data != null)
            {
                MessageBox.Show(string.Format("Непредвиденная ошибка коррекции при записи в строке {0}: {1}", rPos, err.Message));
            }
        }

        private void cmdBack_Click(object sender, EventArgs e)
        {
             UpdateCurrentRowData();

            rPos = rPos - 1;
            if (rPos == 2)
                cmdBack.Enabled = false;
            cmdForward.Enabled = true;

           
            FillCurrentRowData();
        }

        private void cmdForward_Click(object sender, EventArgs e)
        {
            UpdateCurrentRowData();

            rPos = rPos + 1;
            if (rPos == rCount)
                cmdForward.Enabled = false;
            cmdBack.Enabled = true;
            
            FillCurrentRowData();
        }

        private void txtSpell_TextChanged(object sender, EventArgs e)
        {
          
        }
    }
}
