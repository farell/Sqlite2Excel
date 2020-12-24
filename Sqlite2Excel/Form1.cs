using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.IO;
using System.Data.OleDb;

namespace Sqlite2Excel
{
    public partial class Form1 : Form
    {
        private string sourcePath;
        private string outputPath;
        private bool isReallyExit;
        public Form1()
        {
            InitializeComponent();
            this.isReallyExit = false;
        }

        private void buttonSelectSource_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = Application.StartupPath;
            //openFileDialog1.Filter = "sqlite files (*.db)|*.db|All files(*.*)|*>**";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sourcePath = openFileDialog1.FileName;
                textBoxSourceName.Text = sourcePath;
            }
        }

        private void buttonSelectOutput_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = Application.StartupPath;
            saveFileDialog1.Filter = "ext files (*.xls)|*.xls|All files(*.*)|*>**";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == DialogResult.OK && saveFileDialog1.FileName.Length > 0)
            {
                outputPath = saveFileDialog1.FileName;
                textBoxOutputName.Text = outputPath;
            }
        }

        private List<AcquisitionPoint> getOldSensorId()
        {
            SQLiteConnection conn = null;
            List<AcquisitionPoint> ids = new List<AcquisitionPoint>();
            string sql = "select OldSensorId,StakeId,Type from SensorInfo where DtuId in(select DtuId from config where Enabled = 'true')";

            try
            {
                conn = new SQLiteConnection("Data Source ="+sourcePath);
                conn.Open();
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    //Console.WriteLine("Name: " + reader["name"] + "\tScore: " + reader["score"]);
                    //string record = reader.GetString(0) + " " + reader.GetString(1) + "\n";
                    AcquisitionPoint ap = new AcquisitionPoint(reader.GetInt32(0).ToString(), reader.GetString(1), reader.GetString(2));
                    ids.Add(ap);
                    //textBoxLog.AppendText(record);
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                conn.Close();
                //MessageBox.Show(ex.Message);
                Console.WriteLine(ex.ToString());
            }
            return ids;
        }

        private void NewClassify()
        {
            Workbook workbook = new Workbook();

            workbook.Worksheets.Add();
            Worksheet ws = workbook.Worksheets[0];

            Cells cell = ws.Cells;
            cell.SetRowHeight(0, 20);

            ws.Name = "挠度";

            SQLiteConnection conn = null;

            string State = "init";

            string sql = "SELECT Stamp,Value,Type from data where Stamp >'2018-03-08' and SensorId like '85110013%'";

            int i = 0;
            try
            {
                conn = new SQLiteConnection("Data Source =" + sourcePath);
                conn.Open();
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                SQLiteDataReader reader = command.ExecuteReader();

                string topStamp = "2018-03-08 00:07:00";
                //string secondStamp = "2018-03-08 00:07:46";
                DateTime dtRef = Convert.ToDateTime(topStamp);
                while (reader.Read())
                {
                    //Console.WriteLine("Name: " + reader["name"] + "\tScore: " + reader["score"]);
                    //string record = reader.GetString(0) + " " + reader.GetString(1) + "\n";
                    string type = reader.GetString(2);

                    if(type == "offset")
                    {
                        switch (State)
                        {
                            case "init":
                                {
                                    State = "First";
                                    break;
                                }
                            case "First":
                                {
                                    State = "Second";
                                    break;
                                }
                            case "Second":
                                {
                                    //State = "Third";

                                    break;
                                }
                            default:
                                {
                                    State = "init";
                                    break;
                                }
                        }
                    }
                    else
                    {

                    }

                    

                    //TimeSpan ts = dt - dtRef;
                    //if (ts.TotalMinutes < 12)
                    //{
                    //    //时间只差小于10分钟则归为一类
                    //    cell[i, 0].PutValue(topStamp);
                    //    cell[i, 1].PutValue(reader.GetFloat(1));
                    //    dtRef = dt;//更新时间
                    //}
                    //else
                    //{
                    //    topStamp = reader.GetString(0);
                    //    cell[i, 0].PutValue(reader.GetString(0));
                    //    cell[i, 1].PutValue(reader.GetFloat(1));
                    //    dtRef = dt;//更新时间
                    //}
                    i++;
                    //textBoxLog.AppendText(record);
                }
                conn.Close();
                workbook.Save(outputPath);

                MessageBox.Show("导出完成");
            }
            catch (Exception ex)
            {
                conn.Close();
                MessageBox.Show(ex.Message);
            }
        }


        public void ExportDataFromMDB(string date,string fileName)
        {
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sourcePath;
            //string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DSCDdata.mdb";
            string queryString = "SELECT DID,SID,S1,S2,R1, R2,DataTime,IsWarning FROM Data where DataTime > #" + date+"#";

            StreamWriter sw = new StreamWriter(fileName, true);  //true表示如果a.txt文件已存在，则以追加的方式写入

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand(queryString, connection);

                connection.Open();
                OleDbDataReader reader = command.ExecuteReader();

                string topStamp = "2018-03-08 00:07:46";
                //string secondStamp = "2018-03-08 00:07:46";
                DateTime dtRef = Convert.ToDateTime(topStamp); ;
                bool initialized = false;

                while (reader.Read())
                {
                    string did = reader["DID"].ToString();
                    string sensorno = reader["SID"].ToString();
                    string s1String = reader["S1"].ToString();
                    string s2String = reader["S2"].ToString();
                    string r1String = reader["R1"].ToString();
                    string r2String = reader["R2"].ToString();
                    DateTime stamp = (DateTime)reader["DataTime"];
                    string warningString = reader["IsWarning"].ToString();

                    if (!initialized)
                    {
                        //topStamp = Convert.ToDateTime(reader.GetString(1)).ToString("yyyy-MM-dd HH:mm:ss");
                        dtRef = stamp;
                        topStamp = dtRef.ToString("yyyy-MM-dd HH:mm:ss");
                        initialized = true;
                    }

                    TimeSpan ts = stamp - dtRef;
                    if(ts.TotalSeconds > 9*60)
                    {
                        topStamp = stamp.ToString("yyyy-MM-dd HH:mm:ss");
                    }

                    dtRef = stamp;
                    string record = did+"," + sensorno + "," + topStamp + ","+s1String+ ","+s2String+ "," + r1String + "," + r2String + "," +warningString;
                    sw.WriteLine(record);
                }
                sw.Close();
                reader.Close();
            }
            MessageBox.Show("导出完成");
        }

        private void Export2Csv(string fileName)
        {
            StreamWriter sw = new StreamWriter(fileName, true);  //true表示如果a.txt文件已存在，则以追加的方式写入

            SQLiteConnection conn = null;
            //string sql = "SELECT SensorId,Stamp,Type,Value from data where length(SensorId)>5";
            string sql = "SELECT SensorId,Stamp,Type,Value from data";

            int i = 0;
            try
            {
                conn = new SQLiteConnection("Data Source =" + sourcePath);
                conn.Open();
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                SQLiteDataReader reader = command.ExecuteReader();

                string topStamp  = "2018-03-08 00:07:00";
                //string secondStamp = "2018-03-08 00:07:46";
                DateTime dtRef = Convert.ToDateTime(topStamp);
                bool initialized = false;
                while (reader.Read())
                {
                    //Console.WriteLine("Name: " + reader["name"] + "\tScore: " + reader["score"]);
                    //string record = reader.GetString(0) + " " + reader.GetString(1) + "\n";
                    if (!initialized)
                    {
                        topStamp = Convert.ToDateTime(reader.GetString(1)).ToString("yyyy-MM-dd HH:mm:ss");
                        dtRef = Convert.ToDateTime(topStamp);
                        //topStamp = dtRef.ToString("yyyy-MM-dd HH:mm:ss");
                        initialized = true;
                    }
                    string stamp = reader.GetString(1);

                    DateTime dt = Convert.ToDateTime(stamp);

                    TimeSpan ts = dt - dtRef;
                    if (ts.TotalSeconds < 240)
                    {
                        //时间间隔小于5分钟则归为一类
                        string record = reader.GetString(0) + "," + topStamp + "," + reader.GetString(2) + "," + reader.GetFloat(3).ToString();
                        sw.WriteLine(record);
                        //cell[i, 0].PutValue(topStamp);
                        //cell[i, 1].PutValue(reader.GetFloat(1));
                        dtRef = dt;//更新时间
                    }
                    else
                    {
                        topStamp = Convert.ToDateTime(reader.GetString(1)).ToString("yyyy-MM-dd HH:mm:ss");

                        string record = reader.GetString(0) + "," + topStamp + "," + reader.GetString(2) + "," + reader.GetFloat(3).ToString();
                        sw.WriteLine(record);
                        //cell[i, 0].PutValue(reader.GetString(0));
                        //cell[i, 1].PutValue(reader.GetFloat(1));
                        dtRef = dt;//更新时间
                    }
                    i++;
                    //textBoxLog.AppendText(record);
                }
                conn.Close();
                sw.Close();

                MessageBox.Show("导出完成");
            }
            catch (Exception ex)
            {
                sw.Close();
                conn.Close();
                MessageBox.Show(ex.Message);
            }

        }

        private void Classify()
        {
            Workbook workbook = new Workbook();

            workbook.Worksheets.Add();
            Worksheet ws = workbook.Worksheets[0];

            Cells cell = ws.Cells;
            cell.SetRowHeight(0, 20);

            ws.Name = "挠度";

            SQLiteConnection conn = null;

            string sql = "SELECT Stamp,Value from data where Stamp >'2018-03-08' and SensorId like '85110013%'";

            int i = 0;
            try
            {
                conn = new SQLiteConnection("Data Source =" + sourcePath);
                conn.Open();
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                SQLiteDataReader reader = command.ExecuteReader();

                string topStamp = "2018-03-08 00:07:00";
                //string secondStamp = "2018-03-08 00:07:46";
                DateTime dtRef = Convert.ToDateTime(topStamp);
                while (reader.Read())
                {
                    //Console.WriteLine("Name: " + reader["name"] + "\tScore: " + reader["score"]);
                    //string record = reader.GetString(0) + " " + reader.GetString(1) + "\n";
                    //globalStamp = reader.GetString(0);
                    string stamp = reader.GetString(0);
                    topStamp = reader.GetString(0);
                    dtRef = Convert.ToDateTime(topStamp);

                    DateTime dt = Convert.ToDateTime(stamp);
                    //DateTime dt1 = Convert.ToDateTime(globalStamp);
                    TimeSpan ts = dt - dtRef;
                    if (ts.TotalMinutes < 12)
                    {
                        //时间只差小于10分钟则归为一类
                        cell[i, 0].PutValue(topStamp);
                        cell[i, 1].PutValue(reader.GetFloat(1));
                        dtRef = dt;//更新时间
                    }
                    else
                    {
                        topStamp = reader.GetString(0);
                        cell[i, 0].PutValue(reader.GetString(0));
                        cell[i, 1].PutValue(reader.GetFloat(1));
                        dtRef = dt;//更新时间
                    }
                    i++;
                    //textBoxLog.AppendText(record);
                }
                conn.Close();
                workbook.Save(outputPath);

                MessageBox.Show("导出完成");
            }
            catch (Exception ex)
            {
                conn.Close();
                MessageBox.Show(ex.Message);
            }

        }

        //DateTime dt = Convert.ToDateTime("2018-03-09 12:00:00");
        //DateTime dt1 = Convert.ToDateTime("2018-03-08 12:09:00");
        //TimeSpan ts = dt - dt1;
        //MessageBox.Show(ts.TotalSeconds.ToString());

        private void FillWorkSheet(ref Worksheet ws, AcquisitionPoint ap)
        {
            Cells cell = ws.Cells;
            cell.SetRowHeight(0, 20);

            ws.Name = ap.GetId() + "_" + ap.GetStakeId() + "_" + ap.GetDirectionType();

            SQLiteConnection conn = null;

            string sql = "select Stamp,Value from data where SensorId = " + ap.GetId();

            int i = 0;
            try
            {
                conn = new SQLiteConnection("Data Source =" + sourcePath);
                conn.Open();
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    //Console.WriteLine("Name: " + reader["name"] + "\tScore: " + reader["score"]);
                    //string record = reader.GetString(0) + " " + reader.GetString(1) + "\n";
                    cell[i, 0].PutValue(reader.GetString(0));
                    cell[i, 1].PutValue(reader.GetFloat(1));
                    i++;
                    //textBoxLog.AppendText(record);
                }
                conn.Close();
            }
            catch (Exception ex)
            {
                conn.Close();
                MessageBox.Show(ex.Message);
            }

        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(dateTimePicker1.Value.ToShortDateString());
            string dateString = dateTimePicker1.Value.ToShortDateString();
            ExportDataFromMDB(dateString,"out.csv");

            return;

            if (string.IsNullOrEmpty(sourcePath) || string.IsNullOrEmpty(outputPath))
            {
                MessageBox.Show("请选择源文件或输出文件");
                return;
            }

            List<AcquisitionPoint> ids = getOldSensorId();

            if(ids.Count == 0)
            {
                MessageBox.Show("无数据");
                return;
            }

            Workbook workbook = new Workbook();

            for(int i = 0;i< ids.Count;i++)
            {
                workbook.Worksheets.Add();
                Worksheet ws = workbook.Worksheets[i];

                FillWorkSheet(ref ws, ids[i]);
            }

            workbook.Save(outputPath);

            MessageBox.Show("导出完成");

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.WindowState = FormWindowState.Minimized;
                this.notifyIcon1.Visible = true;
                this.Hide();
                if (!this.isReallyExit)
                {
                    this.notifyIcon1.ShowBalloonTip(2000, "提示", "隐藏在任务栏！", ToolTipIcon.Info);
                }
                
                return;
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (this.Visible)
            {
                this.WindowState = FormWindowState.Minimized;
                this.notifyIcon1.Visible = true;
                this.Hide();
            }
            else
            {
                this.Visible = true;
                this.WindowState = FormWindowState.Normal;
                this.Activate();
            }
        }

        private void ToolStripMenuItemRestore_Click(object sender, EventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
            this.notifyIcon1.Visible = true;
            this.Activate();
            //this.Show();
        }

        private void ToolStripMenuItemExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("你确定要退出？", "系统提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                this.isReallyExit = true;
                this.notifyIcon1.Visible = false;
                this.Close();
                //this.Dispose();
                System.Environment.Exit(System.Environment.ExitCode);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string str1 = "2017/2/20 0:31:07";

            DateTime dt = Convert.ToDateTime(str1);

            MessageBox.Show(dt.ToString("yyyy-MM-dd HH:mm:ss"));
        }
    }

}
