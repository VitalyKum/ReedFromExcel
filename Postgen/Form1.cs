using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;


namespace Postgen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //private void Generate()
        //{
        //    string sheetName;
        //    string ConnectionString = String.Format(
        //                    "Provider=Microsoft.ACE.OLEDB.12.0;extended properties=\"excel 8.0;hdr=yes;IMEX=1\";data source={0}",
        //                    "123.xlsx");
        // DataSet ds = new DataSet();
        //    using (OleDbConnection con = new OleDbConnection(ConnectionString))
        //    {
        //        using (OleDbCommand cmd = new OleDbCommand())
        //        {
        //            using (OleDbDataAdapter oda = new OleDbDataAdapter())
        //            {
        //                cmd.Connection = con;
        //                con.Open();
        //                DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        //                for (int i = 0; i < dtExcelSchema.Rows.Count; i++)
        //                {
        //                    sheetName = dtExcelSchema.Rows[i]["TABLE_NAME"].ToString();
        //                    DataTable dt = new DataTable(sheetName);
        //                    cmd.Connection = con;
        //                    //cmd.CommandText = "SELECT SKU as Номер заказа, индекс as Индекс, адрес по русски as Адрес, 收货人手机 as Телефон, Ф.И.О. as ФИО * FROM [" + sheetName + "]";
        //                    cmd.CommandText = "SELECT *  FROM [" + sheetName + "]";
        //                    oda.SelectCommand = cmd;
        //                    oda.Fill(dt);
        //                    dt.Columns.RemoveAt(0);
        //                    dt.Columns[0].ColumnName = "dgbndskjgsrgn";

        //                   dt.TableName = sheetName;
        //                    ds.Tables.Add(dt);
        //                }
        //            }
        //        }
        //    }
        //    ExcelLibrary.DataSetHelper.CreateWorkbook("MyExcelFile.xls", ds);
        //}

        private void cmdChoose_Click(object sender, EventArgs e)
        {
            if (fdoExcel.ShowDialog() == DialogResult.OK)
            {
                txtChoose.Text = fdoExcel.FileName;
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdOk_Click(object sender, EventArgs e)
        {
            string saleID = string.Empty;
            string newsaleID = string.Empty;
            int mass = 0;
            string tel = string.Empty;
            string fio = string.Empty;
            string indx = string.Empty;
            string adr = string.Empty;
            int rCount = 0;

            string code = string.Empty;
            string colorID = string.Empty;

            string gCount = string.Empty;

            string[] goods = new string[] { };
            string[] stringSeparators = new string[] { "\n" };

            txtReport.Text = string.Empty;

            if (txtChoose.Text == "")
            {
                MessageBox.Show("Укажите исходный файл!");
                return;
            }
            
            try
            {
                Excel.Application ex; ex = new Microsoft.Office.Interop.Excel.Application();
                ex.SheetsInNewWorkbook = 1;
                ex.DisplayAlerts = false;
                ex.Visible = false;
                string post_templ = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                post_templ = post_templ + @"\sys\post_templ.xlt";
                string goods_spr = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                goods_spr = goods_spr + @"\sys\goods.xls";

                Excel.Workbook tBook = ex.Workbooks.Open(post_templ);
                Excel.Worksheet tSheet = (Excel.Worksheet)tBook.Worksheets.get_Item(1);

                Excel.Workbook sBook = ex.Workbooks.Open(txtChoose.Text);
                Excel.Worksheet sSheet = (Excel.Worksheet)sBook.Worksheets.get_Item(1);

                Excel.Workbook gBook = ex.Workbooks.Open(goods_spr);
                Excel.Worksheet gSheet = (Excel.Worksheet)gBook.Worksheets.get_Item(1);

                rCount = sSheet.UsedRange.Rows.Count;

                if (rCount < 2)
                {
                    MessageBox.Show("В указанном файле нет данных!");
                    txtChoose.Text = "";
                    return;
                }

                progressBar1.Value = 0;
                progressBar1.Maximum = rCount;

                for (int i = 2; i <= rCount; i++)
                {
                    try
                    {
                        fio = (string)((Excel.Range)sSheet.Cells[i, 6]).Value;
                        if (string.IsNullOrEmpty(fio))
                        {
                            fio = "Нет данных!";
                            txtReport.Text += string.Format("Отсутствуют данные в строке {0}, поле {1}\r\n", i, "ФИО");
                        }
                        ((Excel.Range)tSheet.Cells[i, 2]).Value = BackTranslate(string.Format("{0} {1}", fio, (i-1).ToString()), true);

                        saleID = ((Excel.Range)sSheet.Cells[i, 9]).Value;
                        if (string.IsNullOrEmpty(saleID))
                        {
                            saleID = "Нет данных!";
                            mass = 0;
                            txtReport.Text += string.Format("Отсутствуют данные в строке {0}, поле {1}\r\n", i, "Номер заказа");
                        }
                        else
                        {
                            goods = saleID.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                            mass = 0;
                            newsaleID = string.Empty;
                           
                            foreach (string good in goods)
                            {
                                Goods g = new Goods(good);
                                string searchString = string.Format("{0}{1}", g.Code, g.ColorID);
                                Excel.Range colRange = gSheet.Columns["A:A"];
                                Excel.Range resultRange = colRange.Find(
                                   What: searchString,
                                   LookIn: Excel.XlFindLookIn.xlValues,
                                   LookAt: Excel.XlLookAt.xlPart,
                                   SearchOrder: Excel.XlSearchOrder.xlByRows,
                                   SearchDirection: Excel.XlSearchDirection.xlNext
                                   );
                                if (resultRange != null)
                                {
                                    var v = gSheet.Cells[resultRange.Row, 4].Value;
                                    int t = Convert.ToInt32(v);
                                    mass += t * g.Count;
                                    newsaleID = newsaleID + g.Code + g.ColorID + "-" + g.Count + "\r\n";
                                }                                
                            }
                            if (newsaleID != string.Empty)
                            {
                                newsaleID = newsaleID.Substring(0, newsaleID.Length - 2);
                            }
                            
                            ((Excel.Range)tSheet.Cells[i, 6]).Value = newsaleID == string.Empty ? saleID : newsaleID;

                            double d = Convert.ToDouble(mass) / Convert.ToDouble(1000);
                            //string t_mas = d > 0 ?string.Format("{0}", d) : "0.0";
                            string t_mas = string.Format("{0:F3}", d);
                            //t_mas = t_mas.Replace(",", ".");
                            //((Excel.Range)tSheet.Cells[i, 3]).Value = mass >= 1000 ? mass.ToString() : t_mas;
                            ((Excel.Range)tSheet.Cells[i, 3]).Value = t_mas;
                        }

                        indx = ((Excel.Range)sSheet.Cells[i, 7]).Value;
                        if (string.IsNullOrEmpty(indx))
                        {
                            indx = "Нет данных!";
                            txtReport.Text += string.Format("Отсутствуют данные в строке {0}, поле {1}\r\n", i, "Индекс");
                        }
                        
                        tel = ((Excel.Range)sSheet.Cells[i, 8]).Value;
                        if (string.IsNullOrEmpty(tel))
                        {
                            tel = "Нет данных!";
                            txtReport.Text += string.Format("Отсутствуют данные в строке {0}, поле {1}\r\n", i, "Телефон");
                        }
                        ((Excel.Range)tSheet.Cells[i, 7]).Value = tel;

                        adr = (string)((Excel.Range)sSheet.Cells[i, 3]).Value;
                        if (string.IsNullOrEmpty(adr))
                        {
                            adr = "Нет данных!";
                            txtReport.Text += string.Format("Отсутствуют данные в строке {0}, поле {1}\r\n", i, "Адрес");
                        }
                        else
                        {
                            adr = adr.Replace(indx, "");
                            adr = adr.Replace(tel, "");
                            adr = adr.Replace(fio, "");
                            adr = adr.Replace(", ,", ",");
                            adr = adr.Replace(",,", ",");
                        }
                        ((Excel.Range)tSheet.Cells[i, 1]).Value = BackTranslate(adr, false);
                        
                        ((Excel.Range)tSheet.Cells[i, 8]).Value = string.Format("{0}", 23);
                        ((Excel.Range)tSheet.Cells[i, 10]).Value = string.Format("{0}", 664961);
                    }
                    catch (Exception err) when (err.Data != null)
                    {
                        txtReport.Text += string.Format("Непредвиденная ошибка в строке {0}: {1}", i, err.Message);
                    }
                    progressBar1.Value = i;
                }
                if (chkDoCorrect.Checked)
                {
                    frmCorrect frm = new frmCorrect();
                    frm.tSheet = tSheet;
                    frm.sSheet = sSheet;
                    frm.rCount = rCount - 1;
                   
                    frm.ShowDialog();
                }
                sBook.Close();
                gBook.Close();
                tSheet.Activate();
                ex.DisplayAlerts = true;
                ex.Visible = true;
            }
            catch(Exception err)
            {
                txtReport.Text += string.Format("Ошибка Excel: {0}", err.Message);
            }            
        }

        private string BackTranslate(string st, bool bFio)

        {
            st = st.Replace("kv.", "кв. ");
            st = st.Replace("kv ", "кв. ");
            st = st.Replace("Room ", "кв. ");
            st = st.Replace("room ", "кв. ");
            st = st.Replace("flat ", "кв. ");
            st = st.Replace("Flat ", "кв. ");
            st = st.Replace("kvartira ", "кв. ");
            st = st.Replace("Kvartira ", "кв. ");
            st = st.Replace("street ", "ул. ");
            st = st.Replace("Street ", "ул. ");
            st = st.Replace("ulitsa ", "ул. ");
            st = st.Replace("Ulitsa ", "ул. ");
            st = st.Replace("pereulok", "пер. ");
            st = st.Replace("Pereulok", "пер. ");
            st = st.Replace("dom ", "д. ");
            st = st.Replace("Dom ", "д. ");
            st = st.Replace("house ", "д. ");
            st = st.Replace("House ", "д. ");
            st = st.Replace("prospect", "пр.");
            st = st.Replace("prospekt", "пр.");
            st = st.Replace(" kray ", " кр. ");
            st = st.Replace("Olga ", "Ольга ");
            st = st.Replace("Moscow", "Москва");
            st = st.Replace("moscow", "Москва");
            st = st.Replace("Saint Petersburg", "Санкт-Петербург");
            st = st.Replace("Saint-Petersburg", "Санкт-Петербург"); 

            st = st.Replace("RUSSIA", "");
            st = st.Replace("Russia", "");
            st = st.Replace("russia", "");
            st = st.Replace("oblast", "обл.,");
            st = st.Replace("Oblast", "обл.,");
            st = st.Replace("autonomus", "");
            st = st.Replace("avtonomnyy", "");
            st = st.Replace("autonomous", "");
            st = st.Replace("republic", "респ.");
            st = st.Replace("respublika", "респ.");
            st = st.Replace("okrug", "АО");
            st = st.Replace("schoolroom", "Школьная");
            st = st.Replace("Ty", "Ты");
            st = st.Replace("Sy", "Сы");
            st = st.Replace("Ry", "Ры");
            st = st.Replace("By", "Бы");
            st = st.Replace("Dy", "Ды");

            st = st.Replace("ty", "ты");
            st = st.Replace("sy", "сы");
            st = st.Replace("ry", "ры");
            st = st.Replace("by", "бы");
            st = st.Replace("dy", "ды");

            st = st.Replace("Ch", "Ч");
            st = st.Replace("Sh", "Ш");
            st = st.Replace("Kh", "Х");
            st = st.Replace("H", "Х");
            st = st.Replace("Eh", "Э");
            st = st.Replace("A", "А");
            st = st.Replace("B", "Б");
            st = st.Replace("V", "В");
            st = st.Replace("G", "Г");
            st = st.Replace("D", "Д");
            st = st.Replace("E", "Е");
            st = st.Replace("Jo", "Ё");
            st = st.Replace("Zh", "Ж");
            

            st = st.Replace("Schc", "Щ");
            st = st.Replace("Shh", "Щ");
            st = st.Replace("Sch", "Щ");
            st = st.Replace("Oye", "ое");
            st = st.Replace("Sc", "Ск");
            st = st.Replace("Zw", "Цв");
            st = st.Replace("Yu", "Ю");
            st = st.Replace("Ju", "Ю");
            st = st.Replace("Ya", "Я");
            st = st.Replace("Ch", "Ч");
            st = st.Replace("Sh", "Ш");
            st = st.Replace("Kh", "Х");
            st = st.Replace("H", "Х");
            st = st.Replace("Eh", "Э");
            st = st.Replace("A", "А");
            st = st.Replace("B", "Б");
            st = st.Replace("V", "В");
            st = st.Replace("G", "Г");
            st = st.Replace("D", "Д");
            st = st.Replace("E", "Е");
            st = st.Replace("Jo", "Ё");
            st = st.Replace("Zh", "Ж");
            st = st.Replace("Z", "З");
            st = st.Replace("I", "И");
            st = st.Replace("Y", "Й");
            st = st.Replace("K", "К");
            st = st.Replace("L", "Л");
            st = st.Replace("M", "М");
            st = st.Replace("N", "Н");
            st = st.Replace("O", "О");
            st = st.Replace("P", "П");
            st = st.Replace("R", "Р");
            st = st.Replace("S", "С");
            st = st.Replace("T", "Т");
            st = st.Replace("U", "У");
            st = st.Replace("F", "Ф");
            st = st.Replace("C", "Ц");
            st = st.Replace("'", "Ъ");
            st = st.Replace("I", "Ы");
            st = st.Replace("'", "Ь");
            st = st.Replace("J", "");

            st = st.Replace("schc", "Щ");
            st = st.Replace("shh", "Щ");
            st = st.Replace("sch", "щ");
            st = st.Replace("oye", "ое");
            st = st.Replace("sc", "ск");
            st = st.Replace("Zw", "цв");
            st = st.Replace("yu", "ю");
            st = st.Replace("ju", "ю");
            st = st.Replace("ya", "я");
            st = st.Replace("yy ", "ый ");
            st = st.Replace("yo", "ё");

            st = st.Replace("ye ", "ье ");
            st = st.Replace("ch", "ч");
            st = st.Replace("sh", "ш");
            st = st.Replace("kh", "х");
            st = st.Replace("iy ", "ий ");
            st = st.Replace("ey ", "ей ");
            st = st.Replace("oy ", "ой ");
            st = st.Replace("ay ", "ай ");

            st = st.Replace("ia ", "я ");
            st = st.Replace("jo", "ё");
            st = st.Replace("zh", "ж");
            st = st.Replace("eh", "э");
            st = st.Replace("h", "х");

            if (bFio)
            {
                st = st.Replace("y ", "ий ");                
            }

            st = st.Replace("a", "а");
            st = st.Replace("b", "б");
            st = st.Replace("v", "в");
            st = st.Replace("g", "г");
            st = st.Replace("d", "д");
            st = st.Replace("e", "е");           
            st = st.Replace("z", "з");
            st = st.Replace("i", "и");
            st = st.Replace("y", "й");
            st = st.Replace("k", "к");
            st = st.Replace("l", "л");
            st = st.Replace("m", "м");
            st = st.Replace("n", "н");
            st = st.Replace("o", "о");
            st = st.Replace("p", "п");
            st = st.Replace("r", "р");
            st = st.Replace("s", "с");
            st = st.Replace("t", "т");
            st = st.Replace("u", "у");
            st = st.Replace("f", "ф");
            st = st.Replace("c", "ц");
            st = st.Replace("'", "ъ");
            st = st.Replace("i", "ы");
            st = st.Replace("'", "ь");
            st = st.Replace("j", "");
            

            st = st.Replace("Ци", "Цы");
            st = st.Replace("ци", "цы");
           

            st = st.Replace(",,", ",");
            st = st.Replace(",,", ",");
            st = st.Trim();
            if (st.EndsWith(","))
            {
                st = st.Remove(st.LastIndexOf(","));
            }
            return st; 
        }

        private void cmdSaveReport_Click(object sender, EventArgs e)
        {
            string path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            path = path + @"\report.txt";
            using (StreamWriter sw = new StreamWriter( path, true, System.Text.Encoding.Default))
            {
                sw.WriteLine(string.Format("{0} {1}:\r\n{1}", DateTime.Now.ToString(), txtChoose.Text, txtReport.Text));
                sw.Close();
            }
        }
    }
}
