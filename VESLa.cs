using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using System.Windows;
using Math = System.Math;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using System.Data.OleDb;
using AwokeKnowing.GnuplotCSharp;
using System.Diagnostics;



namespace Win_VES
{
    public partial class Form2 : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel13ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR={1}'";
        public Form2()
        {
            InitializeComponent();
            tabControl1.SelectedIndexChanged += new EventHandler(Tabs_SelectedIndexChanged);
            tabPage2.Enabled = false; // this disables the controls on it
            tabPage3.Enabled = false; // this disables the controls on it
            tabPage4.Enabled = false; // this disables the controls on it
        }
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((Yes.Checked == true) && (No.Checked == false))
            {
                openFileDialog1.ShowDialog();
                // plot data from datagridview
                for (int ir1 = 0; ir1 < (dataGridView1.Rows.Count - 1); ir1++)
                {
                    if (this.dataGridView1.Rows[ir1].Cells[0].Value.ToString() != string.Empty)
                    {
                        if (this.dataGridView1.Rows[ir1].Cells[1].Value.ToString() != string.Empty)
                        {
                            chart1.Series.Clear();
                            chart1.Series.Add("Data Lapangan");
                            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                            {
                                this.chart1.Series["Data Lapangan"].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value.ToString()), Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()));
                            }
                            chart1.ChartAreas[0].AxisX.IsLogarithmic = true;
                            chart1.ChartAreas[0].AxisY.IsLogarithmic = true;
                            chart1.Series["Data Lapangan"].ChartType = SeriesChartType.Line;
                            chart1.ChartAreas[0].AxisX.Title = "AB/2";
                            chart1.ChartAreas[0].AxisY.Title = "RHO";
                        }
                    }
                }
            }
            if ((Yes.Checked == false) && (No.Checked == true))
            {
                MessageBox.Show("Jika data memiliki header silakan cek 'Ya'");
                openFileDialog1.ShowDialog();
                // plot data from datagridview
                for (int ir1 = 0; ir1 < (dataGridView1.Rows.Count - 1); ir1++)
                {
                    if (this.dataGridView1.Rows[ir1].Cells[0].Value.ToString() != string.Empty)
                    {
                        if (this.dataGridView1.Rows[ir1].Cells[1].Value.ToString() != string.Empty)
                        {
                            chart1.Series.Clear();
                            chart1.Series.Add("Data Lapangan");
                            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                            {
                                this.chart1.Series["Data Lapangan"].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value.ToString()), Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()));
                            }
                            chart1.ChartAreas[0].AxisX.IsLogarithmic = true;
                            chart1.ChartAreas[0].AxisY.IsLogarithmic = true;
                            chart1.Series["Data Lapangan"].ChartType = SeriesChartType.Line;
                            chart1.ChartAreas[0].AxisX.Title = "AB/2";
                            chart1.ChartAreas[0].AxisY.Title = "RHO";
                        }
                    }
                }
            }
            if ((Yes.Checked == false) && (No.Checked == false))
            {
                if (MessageBox.Show("Tolong pilih header, Terimakasih!", "Informasi", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    
                }
            }

        }
        private void Tabs_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int ir1 = 0; ir1 < (dataGridView1.Rows.Count - 1); ir1++)
            {
                if (this.dataGridView1.Rows[ir1].Cells[0].Value.ToString() != string.Empty)
                {
                    if (this.dataGridView1.Rows[ir1].Cells[1].Value.ToString() != string.Empty)
                    {

                        if (tabControl1.SelectedTab == tabPage2)
                        {
                            tabPage2.Enabled = true;                            
                            int a4=this.dataGridView1.Rows.Cast<DataGridViewRow>().Select(row => row.Cells[0].Value).Where(value => value != null).Count();
                            textBox1.Text = (a4).ToString();

                            //dataGridView3.Rows.Clear();
                            //if (dataGridView3.ColumnCount == 0)
                            //{
                            //    foreach (DataGridViewColumn dgvc in dataGridView1.Columns)
                            //    {
                            //        dataGridView3.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                            //    }

                            //}
                            //DataGridViewRow row = new DataGridViewRow();
                            //// Adding Rows in datagridview2
                            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            //{
                            //    row = (DataGridViewRow)dataGridView3.Rows[i].Clone();
                            //    int intColIndex = 0;
                            //    foreach (DataGridViewCell cell in dataGridView1.Rows[i].Cells)
                            //    {
                            //        row.Cells[intColIndex].Value = cell.Value;
                            //        intColIndex++;
                            //    }
                            //    dataGridView3.Rows.Add(row);
                            //}
                            for (int ir = 0; ir < (dataGridView1.Rows.Count - 1); ir++)
                            {
                                if (this.dataGridView1.Rows[ir].Cells[0].Value.ToString() != string.Empty)
                                {
                                    if (this.dataGridView1.Rows[ir].Cells[1].Value.ToString() != string.Empty)
                                    {

                                        try
                                        {
                                            chart2.Series.Clear();
                                            chart2.Series.Add("Data Lapangan");
                                            chart2.ResetAutoValues();

                                            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                                            {
                                                this.chart2.Series["Data Lapangan"].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value.ToString()), Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()));

                                            }
                                            chart2.ChartAreas[0].AxisX.IsLogarithmic = true;
                                            chart2.ChartAreas[0].AxisY.IsLogarithmic = true;
                                            chart2.Series[0].ChartType = SeriesChartType.Line;
                                            chart2.ChartAreas[0].AxisX.Title = "AB/2";
                                            chart2.ChartAreas[0].AxisY.Title = "RHO";
                                        }

                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Ada yang salah");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (tabControl1.SelectedTab == tabPage3)
            {
                for (int ir3 = 0; ir3 < (dataGridView2.Rows.Count - 1); ir3++)
                {

                    if (this.dataGridView2.Rows[ir3].Cells[0].Value.ToString() != string.Empty)
                    {
                        if (this.dataGridView2.Rows[ir3].Cells[1].Value.ToString() != string.Empty)
                        {
                            tabPage3.Enabled = true;
                            dataGridView4.Rows.Clear();
                            if (dataGridView4.ColumnCount == 0)
                            {
                                foreach (DataGridViewColumn dgvc in dataGridView2.Columns)
                                {
                                    dataGridView4.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                                }

                            }
                            DataGridViewRow row = new DataGridViewRow();
                            // Adding Rows in datagridview2
                            for (int i = 0; i < dataGridView2.Rows.Count; i++)
                            {
                                row = (DataGridViewRow)dataGridView4.Rows[i].Clone();
                                int intColIndex = 0;
                                foreach (DataGridViewCell cell in dataGridView2.Rows[i].Cells)
                                {
                                    row.Cells[intColIndex].Value = cell.Value;
                                    intColIndex++;
                                }
                                dataGridView4.Rows.Add(row);
                            }
                            //dataGridView5.Rows.Clear();
                            //if (dataGridView5.ColumnCount == 0)
                            //{
                            //    foreach (DataGridViewColumn dgvc in dataGridView3.Columns)
                            //    {
                            //        dataGridView5.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                            //    }

                            //}
                            //DataGridViewRow row1 = new DataGridViewRow();
                            //// Adding Rows in datagridview2
                            //for (int i = 0; i < dataGridView3.Rows.Count; i++)
                            //{
                            //    row = (DataGridViewRow)dataGridView5.Rows[i].Clone();
                            //    int intColIndex = 0;
                            //    foreach (DataGridViewCell cell in dataGridView3.Rows[i].Cells)
                            //    {
                            //        row.Cells[intColIndex].Value = cell.Value;
                            //        intColIndex++;
                            //    }
                            //    dataGridView5.Rows.Add(row);
                            //}

                            try
                            {

                            }

                            catch (Exception ex)
                            {
                                MessageBox.Show("Ada yang salah");
                            }
                        }
                    }
                }


            }


                if (tabControl1.SelectedTab == tabPage4)
                {
                    dataGridView7.Rows.Clear();
                    if (dataGridView7.ColumnCount == 0)
                    {
                        foreach (DataGridViewColumn dgvc in dataGridView4.Columns)
                        {
                            dataGridView7.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                        }

                    }
                    DataGridViewRow row = new DataGridViewRow();
                    // Adding Rows in datagridview2
                    for (int i = 0; i < dataGridView4.Rows.Count; i++)
                    {
                        row = (DataGridViewRow)dataGridView7.Rows[i].Clone();
                        int intColIndex = 0;
                        foreach (DataGridViewCell cell in dataGridView4.Rows[i].Cells)
                        {
                            row.Cells[intColIndex].Value = cell.Value;
                            intColIndex++;
                        }
                        dataGridView7.Rows.Add(row);
                    }

                    dataGridView6.Rows.Clear();
                    if (dataGridView6.ColumnCount == 0)
                    {
                        foreach (DataGridViewColumn dgvc in dataGridView5.Columns)
                        {
                            dataGridView6.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                        }

                    }
                    DataGridViewRow row1 = new DataGridViewRow();
                    // Adding Rows in datagridview2
                    for (int i = 0; i < dataGridView5.Rows.Count; i++)
                    {
                        row = (DataGridViewRow)dataGridView6.Rows[i].Clone();
                        int intColIndex = 0;
                        foreach (DataGridViewCell cell in dataGridView5.Rows[i].Cells)
                        {
                            row.Cells[intColIndex].Value = cell.Value;
                            intColIndex++;
                        }
                        dataGridView6.Rows.Add(row);
                    }
                    try
                    {
                        chart4.Series.Clear();
                        string[] seriesArray = { "s1", "s2", "s3", "s4", "s5", "s6", "s7", "s8", "s9", "s10", "s11", "s12", "s13", "s14", "s15", "s16", "s17", "s18", "s19", "s20" };
                        ArrayList col1Items = new ArrayList();
                        double[] numbers = new double[dataGridView7.Rows.Count];

                        foreach (DataGridViewRow dr in dataGridView7.Rows)
                        {
                            col1Items.Add(dr.Cells[1].Value);
                        }
                        //Graph Series and bind data
                        for (int i = 0; i < dataGridView7.Rows.Count - 1; i++)
                        {
                            numbers[i] = Convert.ToDouble(col1Items[i]);
                            Series series = chart4.Series.Add(seriesArray[i]);
                            series.Points.Add(numbers[i]);
                            chart4.Series[seriesArray[i]].IsValueShownAsLabel = true;
                            series.ChartType = SeriesChartType.StackedColumn;
                            chart4.Series[seriesArray[i]].IsVisibleInLegend = false;
                            ChartArea CA = chart4.ChartAreas[0];
                            CA.AxisY.IsReversed = true;
                            chart4.ChartAreas[0].AxisY.Title = "Kedalaman";
                    }
                    }

                    catch (Exception ex)
                    {

                    }

                }


        }
        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Ingin keluar?", "Keluar", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                this.Close();
            }
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            
                string header = Yes.Checked ? "YES" : "NO";
                string conStr, sheetName;

                conStr = string.Empty;
                switch (extension)
                {

                    case ".xls": //Excel 97-03
                        conStr = string.Format(Excel03ConString, filePath, header);
                        break;

                    case ".xlsx": //Excel 13
                        conStr = string.Format(Excel13ConString, filePath, header);
                        break;

                }

                //Get the name of the First Sheet.
                using (OleDbConnection con = new OleDbConnection(conStr))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        con.Close();
                    }
                }

                //Read Data from the First Sheet.
                using (OleDbConnection con = new OleDbConnection(conStr))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        using (OleDbDataAdapter oda = new OleDbDataAdapter())
                        {
                            DataTable dt = new DataTable();
                            cmd.CommandText = "SELECT * From [" + sheetName + "]";
                            cmd.Connection = con;
                            con.Open();
                            oda.SelectCommand = cmd;
                            oda.Fill(dt);
                            con.Close();

                            //Populate DataGridView.
                            dataGridView1.DataSource = dt;
                        }
                    }
                }
            
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            double[] n1 = new double[19];
            double[] n2 = new double[19];
            double[] a1 = new double[100];
            double[] a2 = new double[100];
            int[] aae = new int[19];
            double[] aaa = new double[100];
            int value;int dw;
            if (Int32.TryParse(textBox1.Text, out value))
            {
                if (Int32.TryParse(textBox3.Text, out dw))
                {
                    if (dw > 1)
                    {
                        for (int i = dataGridView2.Rows.Count - 1; i >= 0; i--)
                        {
                            DataGridViewRow dataGridViewRow = dataGridView2.Rows[i];

                            foreach (DataGridViewCell cell in dataGridViewRow.Cells)
                            {
                                string val = cell.Value as string;
                                if (string.IsNullOrEmpty(val))
                                {
                                    if (!dataGridViewRow.IsNewRow)
                                    {
                                        dataGridView2.Rows.Remove(dataGridViewRow);
                                        break;
                                    }
                                }
                            }
                        }
                        for (int ir1 = 0; ir1 < (dataGridView2.Rows.Count - 1); ir1++)
                        {

                            if (this.dataGridView2.Rows[ir1].Cells[0].Value.ToString() == string.Empty)
                            {
                                if (this.dataGridView2.Rows[ir1].Cells[1].Value.ToString() == string.Empty)
                                {
                                    MessageBox.Show("Input data kedalamann dan resistivitas");
                                    return; // return because we don't want to run normal code of buton click
                                }
                            }
                            else
                            {

                                Int32 index = dataGridView2.Rows.Count - 2;
                                int index1 = Convert.ToInt32(this.dataGridView2.Rows[index].Cells[1].Value.ToString());
                                if (index1 == 0)
                                {
                                    if (value < 1 && value <= 32)
                                    {
                                        MessageBox.Show("Input data point antara 1 dan 32");
                                        return; // return because we don't want to run normal code of buton click
                                    }

                                    else if (textBox3.Text.Trim() == string.Empty)
                                    {
                                        MessageBox.Show("Input jumlah Perlapisan");
                                        return; // return because we don't want to run normal code of buton click
                                    }

                                    else
                                    {
                                        for (int ir = 0; ir < (dataGridView2.Rows.Count - 1); ir++)
                                        {
                                            int a4 = int.Parse(textBox3.Text);
                                            if (this.dataGridView2.Rows.Count < a4)
                                            {

                                                MessageBox.Show("Sesuaikan jumlah baris terhadap jumlah perlapisan");
                                                return; // return because we don't want to run normal code of buton click
                                            }
                                            for (int i = dataGridView2.Rows.Count - 1; i >= 0; i--)
                                            {
                                                DataGridViewRow dataGridViewRow = dataGridView2.Rows[i];

                                                foreach (DataGridViewCell cell in dataGridViewRow.Cells)
                                                {
                                                    string val = cell.Value as string;
                                                    if (string.IsNullOrEmpty(val))
                                                    {
                                                        if (!dataGridViewRow.IsNewRow)
                                                        {
                                                            dataGridView2.Rows.Remove(dataGridViewRow);
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            if (this.dataGridView2.Rows.Count == a4 + 1)
                                            {

                                                n1[ir] = double.Parse(this.dataGridView2.Rows[ir].Cells[0].Value.ToString());
                                                n2[ir] = double.Parse(this.dataGridView2.Rows[ir].Cells[1].Value.ToString());
                                                a1[ir] = n1[ir];
                                                a2[ir] = n2[ir];

                                                //double a3 = 28;// no of data point
                                                //double a5 = 8; // sampling interval
                                                double a3 = int.Parse(textBox1.Text);
                                                double a5 = 8;

                                                double[] b = {0.000318, 0.002072, -0.004978, 0.01125, -0.02521, 0.05812, 0.2494, -1.1324, 2.7044, -3.4507, 0.4248, 1.1817,
                                               0.6194, 0.2374, 0.08688, 0.0235, 0.01284, -0.001198, 0.003042};
                                                double m1 = 2.5 * a5;
                                                double m2 = m1 + a3 - 1;
                                                for (int ir3 = 0; ir3 < (dataGridView1.Rows.Count - 1); ir3++)
                                                {

                                                    if (this.dataGridView1.Rows[ir3].Cells[0].Value.ToString() != string.Empty)
                                                    {
                                                        if (this.dataGridView1.Rows[ir3].Cells[1].Value.ToString() != string.Empty)
                                                        {
                                                            aaa[ir3] = Convert.ToDouble(this.dataGridView1.Rows[ir3].Cells[0].Value.ToString());
                                                            double[] XVAL = aaa;
                                                            double[] n = new double[53];
                                                            double[] L = new double[53];
                                                            double[] T = new double[53];
                                                            double[] temp = new double[53];

                                                            for (int i = 0; i < n.Length; i++)
                                                            {
                                                                n[i] = i;
                                                                L[i] = Math.Pow(10, (2.5556757 - (n[i] / a5)));
                                                                for (int j = a4; j >= 0; j--)
                                                                {
                                                                    if (j == a4)
                                                                    {
                                                                        temp[i] = a1[4];
                                                                    }
                                                                    else
                                                                    {
                                                                        temp[i] = (temp[i] + (a1[j] * (Math.Tanh(L[i] * a2[j])))) / (1 + ((temp[i] * (Math.Tanh(L[i] * a2[j]))) / (a1[j])));
                                                                    }
                                                                }
                                                                T[i] = temp[i];
                                                            }
                                                            double[] TT = new double[100];
                                                            int c = Convert.ToInt32(m1);
                                                            int d = Convert.ToInt32(m2);
                                                            double[] out_data1 = new double[53];
                                                            double[] out_data2 = new double[53];

                                                            for (int a = c - 1; a <= d; a++)
                                                            {
                                                                for (int i = 0; i < b.Length; i++)
                                                                {
                                                                    TT[a] = T[a];
                                                                    out_data1[a] = (XVAL[a - c + 1]);
                                                                    out_data2[a] += b[i] * TT[a];
                                                                }
                                                            }

                                                            //REMOVE NULL ARRAY

                                                            //PLOT DATA to CHART2
                                                            try
                                                            {
                                                                chart2.Series.Clear();
                                                                chart2.Series.Add("Data Perhitungan");
                                                                for (int i = c; i < m2; i++)
                                                                {
                                                                    chart2.Series["Data Perhitungan"].Points.AddXY
                                                                                    (out_data1[i], out_data2[i]);
                                                                }
                                                                chart2.ChartAreas[0].AxisX.IsLogarithmic = true;
                                                                chart2.ChartAreas[0].AxisY.IsLogarithmic = true;
                                                                chart2.Series["Data Perhitungan"].ChartType = SeriesChartType.Line;
                                                                chart2.Series["Data Perhitungan"].Color = Color.Red;
                                                                chart2.Series["Data Perhitungan"].MarkerSize = 10;

                                                                chart2.Series.Add("Data Lapangan");
                                                                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                                                                {
                                                                    this.chart2.Series["Data Lapangan"].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value.ToString()), Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()));

                                                                }
                                                                chart2.Series["Data Lapangan"].ChartType = SeriesChartType.Line;
                                                                chart2.Series["Data Lapangan"].MarkerStyle = MarkerStyle.Circle;
                                                                chart2.Series["Data Lapangan"].Color = Color.Blue;
                                                                chart2.Series["Data Lapangan"].MarkerSize = 10;
                                                            }
                                                            catch (Exception fx)
                                                            {
                                                                MessageBox.Show("Ada yang salah");
                                                            }

                                                            // PLOT DATA to datagridview3
                                                            dataGridView3.Rows.Clear();
                                                            for (int i = c - 1; i <= m2; i++)
                                                            {
                                                                dataGridView3.Rows.Add(new object[] { Math.Round(out_data1[i],3), Math.Round(out_data2[i],3) });
                                                            }
                                                            // SAVE DATA
                                                            System.IO.StreamWriter streamWriter = new System.IO.StreamWriter(
                                                                "D:\\PSE-UGM\\Software-Development\\VES_Processing\\10-Mei\\Win_VES\\demo1.txt");
                                                            string output = "selisih,pangkat ";
                                                            //for (int i = c - 1; i < m2+1; i++)
                                                            //{                                                
                                                            //        output = out_data22[i].ToString() + "\t" + n4[i].ToString();
                                                            //        streamWriter.WriteLine(output);
                                                            //        output = "selisih,pangkat ";
                                                            //}
                                                            for (int i = c - 1; i <= m2; i++)
                                                            {
                                                                output = out_data1[i].ToString() + "\t" + out_data2[i].ToString();
                                                                streamWriter.WriteLine(output);
                                                                output = "selisih,pangkat ";
                                                            }

                                                            streamWriter.Close();
                                                        }
                                                    }
                                                }
                                            }
                                            if (this.dataGridView2.Rows.Count > a4 + 1)
                                            {
                                                MessageBox.Show("Sesuaikan jumlah baris terhadap jumlah perlapisan");
                                                return; // return because we don't want to run normal code of buton click
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Input nilai baris terakhir untuk ketebalan = 0");
                                    return; // return because we don't want to run normal code of buton click
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Input jumlah perlapisan > 1");
                        return; // return because we don't want to run normal code of buton click
                    }
                }

                else
                {
                    MessageBox.Show("Input jumlah perlapisan tanpa koma");
                    return; // return because we don't want to run normal code of buton click
                }
            }
            else
            {
                MessageBox.Show("Input jumlah banyak data tanpa koma");
                return; // return because we don't want to run normal code of buton click
            }
        }        
        private void button2_Click(object sender, EventArgs e)
        {
            double[] n1 = new double[19];
            double[] n2 = new double[19];
            double[] n3 = new double[53];
            double[] n32 = new double[53];
            double[] n4 = new double[53];
            double[] n40 = new double[53];
            double[] a1 = new double[100];
            double[] a11 = new double[100];
            double[] a22 = new double[100];
            double[] rskot = new double[100];
            double[] rsrat = new double[100];
            double[] deprat = new double[100];
            double[] l0 = new double[100];
            double[] l1 = new double[100];
            double[] a2 = new double[100];
            double[] a21 = new double[100];
            double[] depdep = new double[100];
            double[] dep = new double[100];

            if (textBox5.Text.Trim() == string.Empty)
            {
                    MessageBox.Show("Masukkan banyak Iterasi");
                    return; // return because we don't want to run normal code of buton click
            }
           if (double.Parse(textBox5.Text)<1)
           {
                    MessageBox.Show("Masukkan iterasi > 0");
                    return; // return because we don't want to run normal code of buton click
           }
            int value; 
            if (Int32.TryParse(textBox5.Text, out value))
            {
                for (int ir = 0; ir < (dataGridView2.Rows.Count - 1); ir++)
                {
                    if (this.dataGridView2.Rows[ir].Cells[0].Value.ToString() != string.Empty)
                    {
                        if (this.dataGridView2.Rows[ir].Cells[1].Value.ToString() != string.Empty)
                        {
                            //datagridview2=input parameter
                            //datagridview1=data lapangan
                            n1[ir] = double.Parse(this.dataGridView2.Rows[ir].Cells[0].Value.ToString());
                            n2[ir] = double.Parse(this.dataGridView2.Rows[ir].Cells[1].Value.ToString());
                            a1[ir] = n1[ir];//rho
                            a2[ir] = n2[ir];//tebal
                            for (int ir1 = 0; ir1 < (dataGridView1.Rows.Count - 1); ir1++)
                            {
                                if (this.dataGridView1.Rows[ir1].Cells[0].Value.ToString() != string.Empty)
                                {
                                    if (this.dataGridView1.Rows[ir1].Cells[1].Value.ToString() != string.Empty)
                                    {
                                        //double a3 = 28;// no of data point
                                        //double a5 = 8; // sampling interval                                            
                                        n4[ir1] = Convert.ToDouble(dataGridView1.Rows[ir1].Cells[1].Value.ToString());//RHO
                                        n40[ir1] = Convert.ToDouble(dataGridView1.Rows[ir1].Cells[0].Value.ToString());//AB
                                        double a3 = int.Parse(textBox1.Text);
                                        double a5 = 8;
                                        int a4 = int.Parse(textBox3.Text);
                                        double[] out_data1 = new double[53];
                                        double[] out_data2 = new double[53];
                                        double[] out_data3 = new double[53];
                                        double[] out_data4 = new double[53];
                                        double m1 = 2.5 * a5;
                                        double m2 = m1 + a3 - 1;
                                        int c = Convert.ToInt32(m1);
                                        int d = Convert.ToInt32(m2);
                                        PSEVES(a1, a2, a3, a4, a5, ref out_data1, ref out_data2);
                                        //out_data2=app resistivity
                                        //out_data1=AB/2

                                        int aa = 1;
                                        double[] out_data22 = new double[53];
                                        double[] sel_out_data22 = new double[53];
                                        double[] pow_out_data22 = new double[53];
                                        double[] sel_out_data23 = new double[53];
                                        double[] pow_out_data23 = new double[53];
                                        double err0; double err1; double derr; double p;

                                        //PLOT DATA to datagridview5 // tabel AB dan rho
                                        dataGridView5.Rows.Clear();
                                        for (int i = c - 1; i <= m2; i++)
                                        {
                                            dataGridView5.Rows.Add(new object[] { Math.Abs(Math.Round(out_data1[i],3)), Math.Abs(Math.Round(out_data2[i],3)) });
                                        }
                                        
                                        // GET DATA from out data 2 to get error value
                                        for (int i = c - 1; i <= m2; i++)
                                        {
                                            sel_out_data22[i] = ((out_data2[i] - n4[i]));
                                            pow_out_data22[i] = Math.Pow(sel_out_data22[i], 2);
                                        }


                                        err0 = (Mean(pow_out_data22, 0, 50))/1000;
                                        //err0 = pow_out_data22[50];
                                        textBox4.Text = err0.ToString();
                                        double t0 = 1;
                                        double tt = t0;
                                        double maxt = int.Parse(textBox5.Text);
                                        double maxst = 1;
                                        double[] columnData = (from DataGridViewRow row in dataGridView2.Rows
                                                            where row.Cells[0].FormattedValue.ToString() != string.Empty
                                                            select Convert.ToDouble(row.Cells[0].FormattedValue)).ToArray();
                                        double minrho = columnData.Min();
                                        double maxrho = columnData.Max();
                                        double[] columnData1 = (from DataGridViewRow row in dataGridView2.Rows
                                                            where row.Cells[0].FormattedValue.ToString() != string.Empty
                                                            select Convert.ToDouble(row.Cells[1].FormattedValue)).ToArray();
                                        double mindep = columnData1.Min();
                                        double maxdep = columnData1.Max();
                                        double a1l = a1.Length;
                                        double[] m0 = new double[a1.Length];
                                        double[] m00 = new double[a1.Length];
                                        for (int j = 0; j < a1.Length; j++)
                                        {
                                            m0[j] = a1[j];
                                        }
                                        for (int j = 0; j < a1.Length; j++)
                                        {
                                            m00[j] = a2[j];
                                        }

                                        double[] m3 = new double[m0.Length];
                                        double[] m33 = new double[m00.Length];
                                        for (int j = 0; j < a1.Length; j++)
                                        {
                                            m3[j] = m0[j];
                                        }
                                        for (int j = 0; j < a1.Length; j++)
                                        {
                                            m33[j] = m00[j];
                                        }
                                        depdep = LinSpace(0,40, a1.Length);
                                        dep = LinSpace(0, 40, a1.Length);
                                        for (int t = 0; t < maxt; t++)
                                        {
                                            for (int r = 0; r < maxst; r++)
                                            {
                                                for (int j = 0; j < m3.Length; j++)
                                                {

                                                    Random rng = new Random();
                                                    //SAMPAI SINI YAA
                                                    //habis ashar lanjut coding setelah while(true)
                                                    double i = (((a4 - 1) * (GenerateDigit(rng) + 1)));
                                                    int ii = Convert.ToInt32(i);
                                                    int flag = 1;
                                                    n1[ir] = double.Parse(this.dataGridView2.Rows[ir].Cells[0].Value.ToString());//rho
                                                    n2[ir] = double.Parse(this.dataGridView2.Rows[ir].Cells[1].Value.ToString());//tebal
                                                    a1[ir] = n1[ir];
                                                    a2[ir] = n2[ir];
                                                    flag = flag + 1;
                                                    double ui = GenerateDigit(rng);
                                                    double yi = Math.Sign(ui - 0.5) * (tt / t0) * ((Math.Pow((1 + (t0 / tt)), (Math.Abs(2 * ui - 1)))) - 1);
                                                    m0[ii] = m3[ii] + yi * (maxrho - minrho);
                                                    m00[ii] = m33[ii] + yi * (maxdep - mindep);

                                                    if (m0[ii] <= maxrho && m0[ii] >= minrho)
                                                        {
                                                            break;
                                                        }
                                                        if (flag==1)
                                                        {
                                                            return;
                                                        }
                                                }
                                                for (int j = 0; j < a1.Length; j++)
                                                {
                                                    a11[j] = m0[j];
                                                }
                                                for (int j = 0; j < a1.Length; j++)
                                                {
                                                    a22[j] = m00[j];
                                                }
                                                PSEVES(a11, a22, a3, a4, a5, ref out_data3, ref out_data4);
                                                //PLOT DATA to datagridview5
                                                dataGridView5.Rows.Clear();
                                                for (int i = c - 1; i <= m2; i++)
                                                {
                                                    dataGridView5.Rows.Add(new object[] { Math.Abs(Math.Round(out_data3[i],3)), Math.Abs(Math.Round(out_data4[i],3)) });
                                                }
                                                //// GET DATA from datagridview5
                                                //for (int ir2 = 0; ir2 < (dataGridView5.Rows.Count - 1); ir2++)
                                                //{
                                                //    if (this.dataGridView5.Rows[ir2].Cells[0].Value.ToString() != string.Empty)
                                                //    {
                                                //        if (this.dataGridView5.Rows[ir2].Cells[1].Value.ToString() != string.Empty)
                                                //        {
                                                //            //double a3 = 28;// no of data point
                                                //            //double a5 = 8; // sampling interval
                                                //            //n32=rho2
                                                //            n3[ir2] = Convert.ToDouble(dataGridView5.Rows[ir2].Cells[1].Value.ToString());
                                                //            sel_out_data22[ir2] = ((n3[ir2] - n4[ir2]));
                                                //            pow_out_data22[ir2] = Math.Pow(sel_out_data22[ir2], 2);

                                                //        }
                                                //    }
                                                //}
                                                // GET DATA from out data 2 to get error value
                                                for (int i = c - 1; i <= m2; i++)
                                                {
                                                    sel_out_data23[i] = ((out_data4[i] - n4[i]));
                                                    pow_out_data23[i] = Math.Pow(sel_out_data23[i], 2);
                                                }

                                                err1 = (Mean(pow_out_data23, 0, 50))/1000;
                                                derr = err1 - err0;
                                                p = Math.Pow((-derr / tt), 2);
                                                aa = 1;
                                                for (int i = 0; i <= 1; i++)
                                                {
                                                    rskot[aa] = a11[i];
                                                    rskot[aa + 1] = a11[i];
                                                    aa = aa + 2;
                                                }
                                                rskot[4] = rskot[3];
                                                rskot[3] = rskot[2];
                                                for (int k = 0; k < a11.Length; k++)
                                                {
                                                    rsrat[k] = m0[k];
                                                }
                                                for (int k = 2; k < a11.Length-1; k++)
                                                {
                                                    rsrat[k] = Math.Round(((rsrat[k-1] + rsrat[k+1] + rsrat[k]) / 3),2);
                                                }
                                                for (int k = 0; k < a11.Length; k++)
                                                {
                                                    deprat[k] = m00[k];
                                                }
                                                for (int k = 2; k < a11.Length - 1; k++)
                                                {
                                                    deprat[k] =Math.Round(( (deprat[k - 1] + deprat[k + 1] + deprat[k]) / 3),2);
                                                }


                                                if (err1 <= err0)
                                                {
                                                    textBox4.Text = err1.ToString();
                                                    err0 = err1;
                                                    m3 = a11;
                                                    //PLOT DATA to datagridview5
                                                    dataGridView5.Rows.Clear();
                                                    for (int i = c - 1; i <= m2; i++)
                                                    {
                                                        dataGridView5.Rows.Add(new object[] { Math.Abs(Math.Round(out_data3[i], 3)), Math.Abs(Math.Round(out_data4[i], 3)) });
                                                    }
                                                    //PLOT DATA to datagridview4
                                                    dataGridView4.Rows.Clear();
                                                    for (int ir5 = 0; ir5 < (dataGridView2.Rows.Count - 1); ir5++)
                                                    {
                                                        dataGridView4.Rows.Add(new object[] { rsrat[ir5], deprat[ir5] });
                                                    }
                                                        try
                                                    {
                                                        chart3.Series.Clear();
                                                        chart3.Series.Add("Data Perhitungan");
                                                        for (int i = 0; i < (dataGridView5.Rows.Count - 2); i++)
                                                        {
                                                            this.chart3.Series["Data Perhitungan"].Points.AddXY(Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value.ToString()), Convert.ToDouble(dataGridView5.Rows[i].Cells[1].Value.ToString()));

                                                        }
                                                        chart3.ChartAreas[0].AxisX.IsLogarithmic = true;
                                                        chart3.ChartAreas[0].AxisY.IsLogarithmic = true;
                                                        chart3.Series["Data Perhitungan"].ChartType = SeriesChartType.Line;
                                                        chart3.Series["Data Perhitungan"].Color = Color.Red;
                                                        chart3.Series.Add("Data Lapangan");
                                                        for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                                                        {
                                                            this.chart3.Series["Data Lapangan"].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value.ToString()), Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()));

                                                        }
                                                        chart3.Series["Data Lapangan"].ChartType = SeriesChartType.Line;
                                                        chart3.Series["Data Lapangan"].Color = Color.Blue;
                                                        chart3.ChartAreas[0].AxisX.Title = "AB/2";
                                                        chart3.ChartAreas[0].AxisY.Title = "RHO";

                                                    }
                                                    catch (Exception fx)
                                                    {
                                                        MessageBox.Show("Ada yang salah");
                                                    }
                                                    break;


                                                }
                                                else if (err1 > err0)
                                                {
                                                    textBox4.Text = err1.ToString();
                                                    Random rand = new Random();
                                                    double ui = rand.Next(0, 1);
                                                    if (p > ui)
                                                    {
                                                        err0 = err1;
                                                        m3 = a11;
                                                        //PLOT DATA to datagridview5
                                                        dataGridView5.Rows.Clear();
                                                        for (int i = c - 1; i <= m2; i++)
                                                        {
                                                            dataGridView5.Rows.Add(new object[] { Math.Abs(Math.Round(out_data3[i], 3)), Math.Abs(Math.Round(out_data4[i], 3)) });
                                                        }
                                                        //PLOT DATA to datagridview4
                                                        dataGridView4.Rows.Clear();
                                                        for (int ir5 = 0; ir5 < (dataGridView2.Rows.Count - 2); ir5++)
                                                        {
                                                            dataGridView4.Rows.Add(new object[] { rsrat[ir5], deprat[ir5] });
                                                        }
                                                        try
                                                        {

                                                            chart3.Series.Clear();
                                                            chart3.Series.Add("Data Perhitungan");
                                                            for (int i = 0; i < (dataGridView5.Rows.Count - 1); i++)
                                                            {
                                                                this.chart3.Series["Data Perhitungan"].Points.AddXY(Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value.ToString()), Convert.ToDouble(dataGridView5.Rows[i].Cells[1].Value.ToString()));

                                                            }
                                                            chart3.ChartAreas[0].AxisX.IsLogarithmic = true;
                                                            chart3.ChartAreas[0].AxisY.IsLogarithmic = true;
                                                            chart3.Series["Data Perhitungan"].ChartType = SeriesChartType.Line;
                                                            chart3.Series["Data Perhitungan"].Color = Color.Red;
                                                            chart3.Series.Add("Data Lapangan");
                                                            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                                                            {
                                                                this.chart3.Series["Data Lapangan"].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value.ToString()), Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()));

                                                            }
                                                            chart3.Series["Data Lapangan"].ChartType = SeriesChartType.Line;
                                                            chart3.Series["Data Lapangan"].Color = Color.Blue;
                                                            chart3.ChartAreas[0].AxisX.Title = "AB/2";
                                                            chart3.ChartAreas[0].AxisY.Title = "RHO";

                                                        }
                                                        catch (Exception fx)
                                                        {
                                                            MessageBox.Show("Ada yang salah");
                                                        }

                                                    }
                                                    break;
                                                }
                                                else if (err1 < 1)
                                                {
                                                    break;
                                                }


                                            }
                                            
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Input jumlah iterasi tanpa koma");
                return; // return because we don't want to run normal code of buton click
            }

        }        
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            if (this.dataGridView1.Rows.Count == 0)
            {
              MessageBox.Show("Tolong masukkan data", "Informasi",MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);

            }
            for (int ir1 = 0; ir1 < (dataGridView1.Rows.Count - 1); ir1++)
            {
                if (this.dataGridView1.Rows[ir1].Cells[0].Value.ToString() != string.Empty)
                {
                    if (this.dataGridView1.Rows[ir1].Cells[1].Value.ToString() != string.Empty)
                    {


                        SaveFileDialog dialog = new SaveFileDialog();
                        dialog.Filter = "Text File|*.txt";
                        var result = dialog.ShowDialog();
                        if (result != DialogResult.OK)
                            return;
                        StringBuilder builder = new StringBuilder();
                        int rowcount = dataGridView1.Rows.Count;
                        int columncount = dataGridView1.Columns.Count;

                        for (int i = 0; i < rowcount - 1; i++)
                        {
                            List<string> cols = new List<string>();
                            for (int j = 0; j < columncount; j++)
                            {
                                cols.Add(dataGridView1.Rows[i].Cells[j].Value.ToString());
                            }
                            builder.AppendLine(string.Join("\t", cols.ToArray()));
                        }
                        System.IO.File.WriteAllText(dialog.FileName, builder.ToString());
                        MessageBox.Show(@"Text file was created.");
                    }
                }
            }
        }
        private static void PSEVES(double [] a1, double [] a2, double a3, int a4, double a5, ref double [] out_data1, ref double[] out_data2)
        {
            int[] n1 = new int[19];
            int[] n2 = new int[19];

            double[] b = {0.000318, 0.002072, -0.004978, 0.01125, -0.02521, 0.05812, 0.2494, -1.1324, 2.7044, -3.4507, 0.4248, 1.1817,
                                               0.6194, 0.2374, 0.08688, 0.0235, 0.01284, -0.001198, 0.003042};
            double m1 = 2.5 * a5;
            double m2 = m1 + a3 - 1;
            double[] aaa = {1, 1.33352143216332, 1.77827941003892, 2.37137370566166, 3.16227766016838, 4.21696503428582,
                                                5.62341325190349,7.49894209332456, 10, 13.3352143216332, 17.7827941003892, 23.7137370566166, 31.6227766016838,
                                                42.1696503428582, 56.2341325190349, 74.9894209332456, 100, 133.352143216332, 177.827941003892, 237.137370566166,
                                                316.227766016838, 421.696503428582, 562.341325190349, 749.894209332456, 1000, 1333.52143216332, 1778.27941003892,
                                                2371.37370566166, 3162.27766016838, 4216.96503428582, 5623.41325190349,7498.94209332456, 10000};
            double[] XVAL = aaa;
            double[] n = new double[53];
            double[] L = new double[53];
            double[] T = new double[53];
            double[] temp = new double[53];

            for (int i = 0; i < n.Length; i++)
            {
                n[i] = i;
                L[i] = Math.Pow(10, (2.5556757 - (n[i] / a5)));
                for (int j = a4; j >= 0; j--)
                {
                    if (j == a4)
                    {
                        temp[i] = a1[4];
                    }
                    else
                    {
                        temp[i] = (temp[i] + (a1[j] * (Math.Tanh(L[i] * a2[j])))) / (1 + ((temp[i] * (Math.Tanh(L[i] * a2[j]))) / (a1[j])));
                    }
                }
                T[i] = temp[i];
            }
            double[] TT = new double[100];
            int c = Convert.ToInt32(m1);
            int d = Convert.ToInt32(m2);

            for (int a = c - 1; a <= d; a++)
            {
                for (int i = 0; i < b.Length; i++)
                {
                    TT[a] = T[a];
                    out_data1[a] = (XVAL[a - c + 1]);
                    out_data2[a] += b[i] * TT[a];
                }
            }

        }
        public static double Mean(double [] values, int start, int end)
        {
            double s = 0;

            for (int i = start; i < end; i++)
            {
                s += values[i];
            }

            return s / (end - start);
        }
        static double GenerateDigit(Random rng)
        {
            // Assume there'd be more logic here really
            return rng.Next(2);
        }
        public static double[] LinSpace (double x1, double x2, int n)
        {
            double step = (x2 - x1) / (n - 1);
            double[] y = new double[n];
            for (int i=0; i<n; i++)
                y[i] = x1 + step * i;
            return y;            
        }    

        //private int[] buildTeamAData()
        //{
        //    int[] goalsScored = new int[1000];
        //    for (int i = 0; i < 1000; i++)
        //    {
        //        goalsScored[i] = (i + 1000)*11;
        //    }
        //    return goalsScored;
        //}

        //private void plotGraph()
        //{
        //    GraphPane myPane = zedGraphControl1.GraphPane;

        //    // Set the Titles
        //    myPane.Title.Text = "Team A Analysis for 2014/2015 Season";
        //    myPane.XAxis.Title.Text = "Year";
        //    myPane.YAxis.Title.Text = "No of Goals";
        //    myPane.XAxis.Type = ZedGraph.AxisType.Log;
        //    myPane.YAxis.Type = ZedGraph.AxisType.Log;
        //    PointPairList teamAPairList = new PointPairList();
        //    int[] teamAData = buildTeamAData();
        //    for (int i = 0; i < 1000; i++)
        //    {
        //        teamAPairList.Add(teamAData[i], teamAData[i]);
        //    }
        //    LineItem teamACurve = myPane.AddCurve("Team A",teamAPairList, Color.Red, SymbolType.Diamond);

        //    zedGraphControl1.AxisChange();
        //}

        //private void SetSize()
        //{
        //    zedGraphControl1.Location = new Point(0, 0);
        //    zedGraphControl1.IsShowPointValues = true;
        //    zedGraphControl1.Size = new Size(this.ClientRectangle.Width - 20, this.ClientRectangle.Height - 50);

        //}
        //    // SAVE DATA
        //    System.IO.StreamWriter streamWriter = new System.IO.StreamWriter(
        //        "D:\\PSE-UGM\\Software-Development\\VES_Processing\\10-Mei\\Win_VES\\demo1.txt");
        //    string output = "selisih,pangkat ";
        //                                        //for (int i = c - 1; i < m2+1; i++)
        //                                        //{                                                
        //                                        //        output = out_data22[i].ToString() + "\t" + n4[i].ToString();
        //                                        //        streamWriter.WriteLine(output);
        //                                        //        output = "selisih,pangkat ";
        //                                        //}
        //                                        for (int i = 0; i< (dataGridView1.Rows.Count - 1); i++)
        //                                        {
        //                                            output = n4[i].ToString() + "\t" + n4[i].ToString();
        //    streamWriter.WriteLine(output);
        //                                            output = "selisih,pangkat ";
        //                                        }

        //streamWriter.Close();
    }
}
