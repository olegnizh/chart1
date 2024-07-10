/*
 * Created by SharpDevelop.
 * User: alexpc
 * Date: 24.02.2023
 * Time: 14:10
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Configuration;
using System.IO;


namespace chart1
{
	/// <summary>
	/// Description of Form1.
	/// </summary>
	public partial class Form1 : Form
	{
        string[] orgs = null;
        decimal[] vorg = null;
        decimal[] dorg = null;
        decimal max_vorg = 0;
        decimal max_dorg = 0;
        int _dx = 0;
        int _nofstep = 0;


        public Form1()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			Info.path_app = AppDomain.CurrentDomain.BaseDirectory;
            this.dataGridView1.DataSource = Info.dt;
            this.textBox1.Text = "";
            //MessageBox.Show(Info.path_app);
            this.radioButton2.Checked = true;
            //

            Info.s_browser = ConfigurationManager.AppSettings["s_browser"];
            //MessageBox.Show(Info.s_browser.ToString());
            Info.s_work = ConfigurationManager.AppSettings["s_work"];
            //

            this.textBox2.Text = Info.s_work;
            Info.s_work_a = ConfigurationManager.AppSettings["s_work_a"];
            
		}
		
        System.Data.DataTable get_dt()
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                //conn.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" + Info.path_data + "; Extended Properties='Excel 8.0'";                
                conn.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" + Info.s_work + Info.path_data + ".xls; Extended Properties='Excel 8.0;HDR=YES'";
                //conn.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" + path + "\\test.xls; Extended Properties='Excel 8.0'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    //comm.CommandText = "Select [Организация] from [sheet1$];";
                    comm.CommandText = "Select * from [sh$];";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }
 
		void Button1Click(object sender, EventArgs e)
		{
            //
            this.write_html();
            
		}

        void calc_dop()
        {           
            decimal x_max = 0;
            x_max = Convert.ToDecimal(this.t_maxx.Text.ToString().Trim());
            decimal x_dop = 0;
            decimal x = 0;
            
            // calc dop             
            foreach (DataRow r in Info.dt.Rows)
            {
                x = decimal.Round(Convert.ToDecimal(r[1].ToString()));
                x_dop = decimal.Round(x * x_max / this.max_vorg);
                r[2] = x_dop;
                // set max dorg
                if (this.max_dorg <= x_dop)
                    this.max_dorg = x_dop;
            }
        }


        private void button5_Click(object sender, EventArgs e)
        {
            // validate - calc                                   
            if ((this.t_maxx.Text.ToString().Trim() == "") || (this.t_maxx.Text.ToString().Trim() == "0"))
            {
                MessageBox.Show("Параметр (Max X) не корректен - нельзя показать шкалу Х");
                return;
            }
            if ((this.t_dx.Text.ToString().Trim() == "") || (this.t_dx.Text.ToString().Trim() == "0"))
            {
                MessageBox.Show("Параметр (d X) не корректен - нельзя показать шкалу Х");
                return;
            }
            //if ((this.t_count.Text.ToString().Trim() == "") || (this.t_count.Text.ToString().Trim() == "0"))
            //{
            //    MessageBox.Show("Параметр (Количество шагов) не корректен - нельзя показать шкалу Х");
            //    return;
            //}
            // показатели для показа
            this.calc_dop();

            // ширина шага для показа
            string s = "";
            this._dx = 0;
            s = Convert.ToString(this.max_vorg * Convert.ToInt32(this.t_dx.Text.ToString().Trim()) / Convert.ToInt32(this.t_maxx.Text.ToString().Trim()));
            //MessageBox.Show("_dx = "+s);
            //MessageBox.Show(this._dx.ToString());
            //return;
            int pos = s.LastIndexOf(',');
            if (pos != -1)
            {
                //s = s.Substring(0, s.LastIndexOf(','));
                s = s.Remove(s.LastIndexOf(","));
            }
            this._dx = 0;
            this._dx = Convert.ToInt32(s);
            //MessageBox.Show(this._dx.ToString());

            // количество шагов для показа
            //this._nofstep = Convert.ToInt32(this.t_count.Text.ToString().Trim());
            s = Convert.ToString(Convert.ToInt32(this.t_maxx.Text.ToString().Trim()) / Convert.ToInt32(this.t_dx.Text.ToString().Trim()));
            this._nofstep = 0;
            this._nofstep = Convert.ToInt32(s);
            this.t_tnstep.Text = Convert.ToString(this._nofstep);
            //MessageBox.Show(_nofstep.ToString());

        }
        void write_html()
        {
            string s = Info.s_work + Info.s_name + ".html";
            StreamWriter sw = new StreamWriter(s, false);

            orgs = new string[this.dataGridView1.Rows.Count];
            vorg = new decimal[this.dataGridView1.Rows.Count];
            dorg = new decimal[this.dataGridView1.Rows.Count];

            decimal max_width = 0;
            int init_y_bar = 1;
            int max_height = 0;
            int bar_height = 19;
            for (int i = 0; i < this.dataGridView1.Rows.Count; ++i)
            {
                orgs[i] = this.dataGridView1[0, i].Value.ToString();
                vorg[i] = decimal.Round(Convert.ToDecimal(this.dataGridView1[1, i].Value.ToString()));
                dorg[i] = decimal.Round(Convert.ToDecimal(this.dataGridView1[2, i].Value.ToString()));
                //MssageBox.Show(orgs[i]+"  "+ vorg[i].ToString());                
            }
            max_width = this.max_dorg + 500;
            max_height = init_y_bar + (19 + 10) * this.dataGridView1.Rows.Count + 50;

            sw.Write(@"<!DOCTYPE html>
<html>
<head>
  <title>" + Info.s_name + "</title>");
            sw.Write(sw.NewLine);
            sw.Write(@"  <style>
.navbar {
    overflow: hidden;
    background-color: white;
    position: fixed;
    top: 0;
    width: 100%;
}
.main {
  margin-top: 30px;
}
.bar {
  fill: red;
  height: 21px;
  transition: fill .3s ease;
  cursor: pointer;
  font-family: Helvetica, sans-serif;
}
.bar:hover,
.bar:focus {
  fill: black;
}
.bar:hover text,
.bar:focus text {
  fill: red;
}  
  </style>
</head>
<body>");
            sw.Write(sw.NewLine);
            sw.Write(@"<div class='navbar'>");
            sw.Write(sw.NewLine);
            // begin шкала X
            sw.Write(@"<svg height = '25' width = '"+ max_width.ToString() + "'>");
            sw.Write(sw.NewLine);
            sw.Write(@"<line x1 = '0' y1 = '0' x2 = '" + this.max_dorg.ToString() + "' y2 = '0' style = 'stroke:rgb(0,0,0);stroke-width:2'/>");
            sw.Write(sw.NewLine);
            // Min X
            sw.Write(@"<line x1 = '0' y1 = '0' x2 = '0' y2 = '10' style = 'stroke:rgb(0,0,0);stroke-width:2'/>");
            sw.Write(sw.NewLine);
            // Max X
            sw.Write(@"<line x1 = '"+ this.max_dorg.ToString() + "' y1 = '0' x2 = '"+ this.max_dorg.ToString() + "' y2 = '10' style = 'stroke:rgb(0,0,0);stroke-width:2'/>");
            sw.Write(sw.NewLine);
            sw.Write(@"<text x = '" + (this.max_dorg + 3).ToString() + "' y = '23' fill = 'black'>Max "+ this.max_vorg.ToString() + "</text>");
            sw.Write(sw.NewLine);
            // 
            int x = 0;
            int step = Convert.ToInt32(this.t_dx.Text.ToString().Trim());
            for (int m = 1; m < this._nofstep; ++m)
            {
                x = x + step;
                sw.Write(@"<line x1 = '" + x.ToString() + "' y1 = '0' x2 = '" + x.ToString() + "' y2 = '10' style = 'stroke:rgb(0,0,0);stroke-width:2'/>");
                sw.Write(sw.NewLine);
                sw.Write(@"<text x = '" + (x - 5).ToString() + "' y = '23' fill = 'black'>" + (m * this._dx).ToString() + "</text>");
                sw.Write(sw.NewLine);
            }
            // 
            sw.Write(@"</svg>");
            sw.Write(sw.NewLine);
            sw.Write(@"</div>");
            sw.Write(sw.NewLine);

            //sw.Write(@"<br>");
            //sw.Write(sw.NewLine);
            // end шкала X

            sw.Write(@"<div class='main'>");
            sw.Write(sw.NewLine);
            sw.Write(@"<svg class='chart' width='" + max_width.ToString() + "' height='" + max_height.ToString() + "' aria-labelledby='title desc' role='img'>");
            sw.Write(sw.NewLine);

            // j == 0
            sw.Write(@"<g class='bar'>");
            sw.Write(sw.NewLine);
            sw.Write(@"<rect width='" + dorg[0].ToString() + "' height='" + bar_height.ToString() + "' y='" + init_y_bar.ToString() + "'></rect>");
            sw.Write(sw.NewLine);
            sw.Write(@"<text x='" + (dorg[0] + 3).ToString() + "' y='"+ (init_y_bar + 10).ToString() + "' dy='.35em' fill='black'>" + orgs[0].ToString() + "</text>");
            sw.Write(sw.NewLine);
            sw.Write(@"</g>");
            sw.Write(sw.NewLine);
            //
            for (int j = 1; j < this.dataGridView1.Rows.Count; ++j)
            {
                sw.Write(@"<g class='bar'>");
                sw.Write(sw.NewLine);
                sw.Write(@"<rect width='" + dorg[j].ToString() + "' height='" + bar_height.ToString() + "' y='" + (init_y_bar + 20 * j).ToString() + "'></rect>");
                sw.Write(sw.NewLine);
                sw.Write(@"<text x='" + (dorg[j] + 3).ToString() + "' y='" + (init_y_bar + 10 + 20 * j).ToString() + "' dy='.35em' fill='black'>" + orgs[j].ToString() + "</text>");
                sw.Write(sw.NewLine);
                sw.Write(@"</g>");
                sw.Write(sw.NewLine);
            }
            sw.Write(@"</svg>");
            sw.Write(sw.NewLine);
            sw.Write(@"</div>");
            sw.Write(sw.NewLine);
            sw.Write(@"</body>");
            sw.Write(sw.NewLine);
            sw.Write(@"</html>");           

            sw.Close();

            MessageBox.Show("Файл html создан : " + Info.s_work + Info.s_name + ".html");

        }


        void Button2Click(object sender, EventArgs e)
		{
            // visualizer
            //Form f2 = new Form2();            
            //f2.ShowDialog();
            /*
            Process proc = new ();
            proc.StartInfo.FileName = "C:\\Program Files\\Mozilla Firefox\\firefox.exe";
            //proc.StartInfo.Arguments = "<command arguments>";
            proc.Start();
            */
            //Process.Start("C:\\Program Files\\Mozilla Firefox\\firefox.exe", "C:\\temp\\bar-vert-1.html");

            // встроенный
            if (this.radioButton1.Checked)
                Info.s_browser = Info.path_app + "FirefoxPortable64\\FirefoxPortable.exe";
            // текущий
            if (this.radioButton2.Checked)
                Info.s_browser = "";
            // из config
            if (this.radioButton3.Checked)
                Info.s_browser = ConfigurationManager.AppSettings["s_browser"];

            //MessageBox.Show(Info.s_browser +" : "+ Info.s_work + Info.s_name + ".html");
            if (Info.s_browser == "")
                Process.Start(Info.s_work + Info.s_name + ".html");
            else
                Process.Start(Info.s_browser, Info.s_work + Info.s_name + ".html");
            
            
		}

        private void button3_Click(object sender, EventArgs e)
        {
            // open file
            OpenFileDialog d = new OpenFileDialog();
            d.InitialDirectory = Info.s_work;
            d.Title = "Выберите XLS файл";
            d.Filter = "Файл Excel (*.xls)|*.xls";
            if (d.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = Path.GetFileNameWithoutExtension(d.FileName);
                Info.path_data = this.textBox1.Text;
                Info.s_name = this.textBox1.Text;

            }
            else
                return;

            // get excel data
            if (Info.dt != null)
            {
                Info.dt.Dispose();
                Info.dt = null;
            }
            Info.dt = get_dt();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.DataSource = Info.dt;
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            // init max_vorg
            decimal x = 0;
            foreach (DataRow r in Info.dt.Rows)
            {
                x = decimal.Round(Convert.ToDecimal(r[1].ToString()));
                if (this.max_vorg <= x)
                    this.max_vorg = x;
            }
            this.l_maxvorg.Text = "Max показатель : " + this.max_vorg.ToString();


        }

        private void button4_Click(object sender, EventArgs e)
        {
            //string msg = String.Format("Row: {0}, Column: {1}", this.dataGridView1.CurrentCell.RowIndex, this.dataGridView1.CurrentCell.ColumnIndex);
            //MessageBox.Show(msg, "Current Cell");

            //MessageBox.Show(this.dataGridView1[this.dataGridView1.CurrentCell.ColumnIndex, this.dataGridView1.CurrentCell.RowIndex].Value.ToString());

 



        }



    }
}
