using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;

namespace K_Means_Clustering
{
    public partial class Form1 : Form
    {
        //Public Variables
        string fileName;
        int k;
        double centerX1;
        double centerY1;
        double centerX2;
        double centerY2;
        double centerX3;
        double centerY3;
        List<double> newCenterX1 = new List<double>();
        List<double> newCenterY1 = new List<double>();
        List<double> newCenterX2 = new List<double>();
        List<double> newCenterY2 = new List<double>();
        List<double> newCenterX3 = new List<double>();
        List<double> newCenterY3 = new List<double>();
        double newCenterPointX1 = 0;
        double newCenterPointY1 = 0;
        double newCenterPointX2 = 0;
        double newCenterPointY2 = 0;
        double newCenterPointX3 = 0;
        double newCenterPointY3 = 0;
        int counter = 0;
        bool repeat = true;
        int digitNumberAfterDot;
        int maxRepeat;
        bool runOnce = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            repeat = true;
            k = Convert.ToInt32(numericUpDown1.Value);
            digitNumberAfterDot = Convert.ToInt32(numericUpDown2.Value);
            maxRepeat = Convert.ToInt32(numericUpDown3.Value);
            dataGridView2.Rows.Clear();
            dataGridView2.Refresh();

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel WorkBook 97-2003|*.xls" })
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    this.toolStripStatusLabel1.Text = fileName;
                    toolStripStatusLabel1.Text = "Excel file loaded.";
                }
            }
            
            string PathConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties = \"Excel 12.0 Xml;HDR=YES\"; ";
            OleDbConnection conn = new OleDbConnection(PathConn);
            try
            {
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from [" + "Sayfa1" + "$]", conn);
                System.Data.DataTable dt = new System.Data.DataTable();
                myDataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
                toolStripStatusLabel1.Text = "[TR] File is on view. Ready to run.";
            }
            catch (Exception)
            {
                try
                {
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from [" + "Sheet1" + "$]", conn);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    myDataAdapter.Fill(dt);
                    dataGridView1.DataSource = dt;
                    toolStripStatusLabel1.Text = "[EN] File is on view. Ready to run.";
                }
                catch (Exception)
                {
                    toolStripStatusLabel1.Text = "Wrong language or file format!";
                }
            }
        }

        private double distanceToCenterFunction(double centerX, double centerY, double pointX, double pointY)
        {
            var result = Math.Sqrt(Math.Pow(centerX - pointX, 2) + Math.Pow(centerY - pointY, 2));
            return result;
        }

        private void runToolStripMenuItem_Click(object sender, EventArgs e)
        {
            centerX1 = Convert.ToDouble(dataGridView1.Rows[0].Cells[1].Value);
            centerY1 = Convert.ToDouble(dataGridView1.Rows[0].Cells[2].Value);
            centerX2 = Convert.ToDouble(dataGridView1.Rows[2].Cells[1].Value);
            centerY2 = Convert.ToDouble(dataGridView1.Rows[2].Cells[2].Value);
            centerX3 = Convert.ToDouble(dataGridView1.Rows[3].Cells[1].Value);
            centerY3 = Convert.ToDouble(dataGridView1.Rows[3].Cells[2].Value);
            int rowLenght = Convert.ToInt32(dataGridView1.Rows.GetRowCount(0)) - 1;

            if (k == 2)
            {
                while (repeat == true)
                {
                    for (int i = 0; i < rowLenght; i++)
                    {
                        double pointX = Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value);
                        double pointY = Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value);
                        double centerA = distanceToCenterFunction(centerX1, centerY1, pointX, pointY);
                        double centerB = distanceToCenterFunction(centerX2, centerY2, pointX, pointY);

                        if (centerA < centerB)
                        {
                            newCenterX1.Add(pointX);
                            newCenterY1.Add(pointY);
                        }
                        else
                        {
                            newCenterX2.Add(pointX);
                            newCenterY2.Add(pointY);
                        }
                    }
                    for (int i = 0; i < newCenterX1.Count; i++)
                    {
                        newCenterPointX1 = newCenterPointX1 + newCenterX1[i];
                    }
                    for (int i = 0; i < newCenterY1.Count; i++)
                    {
                        newCenterPointY1 = newCenterPointY1 + newCenterY1[i];
                    }
                    for (int i = 0; i < newCenterX2.Count; i++)
                    {
                        newCenterPointX2 = newCenterPointX2 + newCenterX2[i];
                    }
                    for (int i = 0; i < newCenterY2.Count; i++)
                    {
                        newCenterPointY2 = newCenterPointY2 + newCenterY2[i];
                    }
                    newCenterPointX1 = newCenterPointX1 / newCenterX1.Count;
                    newCenterPointY1 = newCenterPointY1 / newCenterY1.Count;
                    newCenterPointX2 = newCenterPointX2 / newCenterX2.Count;
                    newCenterPointY2 = newCenterPointY2 / newCenterY2.Count;

                    centerX1 = Math.Round(centerX1, digitNumberAfterDot);
                    centerY1 = Math.Round(centerY1, digitNumberAfterDot);
                    centerX2 = Math.Round(centerX2, digitNumberAfterDot);
                    centerY2 = Math.Round(centerY2, digitNumberAfterDot);
                    newCenterPointX1 = Math.Round(newCenterPointX1, digitNumberAfterDot);
                    newCenterPointY1 = Math.Round(newCenterPointY1, digitNumberAfterDot);
                    newCenterPointX2 = Math.Round(newCenterPointX2, digitNumberAfterDot);
                    newCenterPointY2 = Math.Round(newCenterPointY2, digitNumberAfterDot);

                    if (centerX1 == newCenterPointX1 && centerY1 == newCenterPointY1 && centerX2 == newCenterPointX2 && centerY2 == newCenterPointY2 || counter == maxRepeat)
                    {
                        toolStripStatusLabel1.Text = "Process Complete.";
                        if (counter == maxRepeat)
                        {
                            toolStripStatusLabel1.Text = "Program stopped because the algorithm was repeated" + " " + maxRepeat + " " + "times.";
                        }

                        if (runOnce == false)
                        {
                            dataGridView2.Columns.Add("ColunmName", " ");
                            dataGridView2.Columns.Add("ColunmName", "X");
                            dataGridView2.Columns.Add("ColunmName", "Y");
                            runOnce = true;
                        }

                        for (int i = 0; i < k; i++)
                        {
                            dataGridView2.Rows.Add();
                            dataGridView2.Rows[i].Cells[0].Value = "Cluster Center" + " " + (i+1);
                        }
                        dataGridView2.Rows[0].Cells[1].Value = newCenterPointX1;
                        dataGridView2.Rows[0].Cells[2].Value = newCenterPointY1;
                        dataGridView2.Rows[1].Cells[1].Value = newCenterPointX2;
                        dataGridView2.Rows[1].Cells[2].Value = newCenterPointY2;

                        repeat = false;
                    }
                    else
                    {
                        repeat = true;
                        counter = counter + 1;
                    }
                    centerX1 = newCenterPointX1;
                    centerY1 = newCenterPointY1;
                    centerX2 = newCenterPointX2;
                    centerY2 = newCenterPointY2;
                }
            }
            else if (k == 3)
            {
                while (repeat == true)
                {
                    for (int i = 0; i < rowLenght; i++)
                    {
                        double pointX = Convert.ToInt32(dataGridView1.Rows[i].Cells[1].Value);
                        double pointY = Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value);
                        double distanceToCenter1 = distanceToCenterFunction(centerX1, centerY1, pointX, pointY);
                        double distanceToCenter2 = distanceToCenterFunction(centerX2, centerY2, pointX, pointY);
                        double distanceToCenter3 = distanceToCenterFunction(centerX3, centerY3, pointX, pointY);

                        if (distanceToCenter1 < distanceToCenter2 && distanceToCenter1 < distanceToCenter3)
                        {
                            newCenterX1.Add(pointX);
                            newCenterY1.Add(pointY);
                        }
                        else if (distanceToCenter2 < distanceToCenter1 && distanceToCenter2 < distanceToCenter3)
                        {
                            newCenterX2.Add(pointX);
                            newCenterY2.Add(pointY);
                        }
                        else if (distanceToCenter3 < distanceToCenter1 && distanceToCenter3 < distanceToCenter2)
                        {
                            newCenterX3.Add(pointX);
                            newCenterY3.Add(pointY);
                        }
                        else
                        {
                            MessageBox.Show("Error Code: distanceToCenter ");
                        }
                    }
                    for (int i = 0; i < newCenterX1.Count; i++)
                    {
                        newCenterPointX1 = newCenterPointX1 + newCenterX1[i];
                    }
                    for (int i = 0; i < newCenterY1.Count; i++)
                    {
                        newCenterPointY1 = newCenterPointY1 + newCenterY1[i];
                    }
                    for (int i = 0; i < newCenterX2.Count; i++)
                    {
                        newCenterPointX2 = newCenterPointX2 + newCenterX2[i];
                    }
                    for (int i = 0; i < newCenterY2.Count; i++)
                    {
                        newCenterPointY2 = newCenterPointY2 + newCenterY2[i];
                    }
                    for (int i = 0; i < newCenterX3.Count; i++)
                    {
                        newCenterPointX3 = newCenterPointX3 + newCenterX3[i];
                    }
                    for (int i = 0; i < newCenterY3.Count; i++)
                    {
                        newCenterPointY3 = newCenterPointY3 + newCenterY3[i];
                    }
                    newCenterPointX1 = newCenterPointX1 / newCenterX1.Count;
                    newCenterPointY1 = newCenterPointY1 / newCenterY1.Count;
                    newCenterPointX2 = newCenterPointX2 / newCenterX2.Count;
                    newCenterPointY2 = newCenterPointY2 / newCenterY2.Count;
                    newCenterPointX3 = newCenterPointX3 / newCenterX3.Count;
                    newCenterPointY3 = newCenterPointY3 / newCenterY3.Count;

                    centerX1 = Math.Round(centerX1, digitNumberAfterDot);
                    centerY1 = Math.Round(centerY1, digitNumberAfterDot);
                    centerX2 = Math.Round(centerX2, digitNumberAfterDot);
                    centerY2 = Math.Round(centerY2, digitNumberAfterDot);
                    centerX3 = Math.Round(centerX3, digitNumberAfterDot);
                    centerY3 = Math.Round(centerY3, digitNumberAfterDot);
                    newCenterPointX1 = Math.Round(newCenterPointX1, digitNumberAfterDot);
                    newCenterPointY1 = Math.Round(newCenterPointY1, digitNumberAfterDot);
                    newCenterPointX2 = Math.Round(newCenterPointX2, digitNumberAfterDot);
                    newCenterPointY2 = Math.Round(newCenterPointY2, digitNumberAfterDot);
                    newCenterPointX3 = Math.Round(newCenterPointX3, digitNumberAfterDot);
                    newCenterPointY3 = Math.Round(newCenterPointY3, digitNumberAfterDot);

                    if (centerX1 == newCenterPointX1 && centerY1 == newCenterPointY1 && centerX2 == newCenterPointX2 && centerY2 == newCenterPointY2 && centerX3 == newCenterPointX3 && centerY3 == newCenterPointY3 || counter == maxRepeat)
                    {
                        toolStripStatusLabel1.Text = "Process Complete.";
                        if (counter == maxRepeat)
                        {
                            toolStripStatusLabel1.Text = "Program stopped because the algorithm was repeated" + " " + maxRepeat + " " + "times.";
                        }

                        if (runOnce == false)
                        {
                            dataGridView2.Columns.Add("ColunmName", " ");
                            dataGridView2.Columns.Add("ColunmName", "X");
                            dataGridView2.Columns.Add("ColunmName", "Y");
                            runOnce = true;
                        }

                        for (int i = 0; i < k; i++)
                        {
                            dataGridView2.Rows.Add();
                            dataGridView2.Rows[i].Cells[0].Value = "Cluster Center" + " " + (i + 1);
                        }
                        dataGridView2.Rows[0].Cells[1].Value = newCenterPointX1;
                        dataGridView2.Rows[0].Cells[2].Value = newCenterPointY1;
                        dataGridView2.Rows[1].Cells[1].Value = newCenterPointX2;
                        dataGridView2.Rows[1].Cells[2].Value = newCenterPointY2;
                        dataGridView2.Rows[2].Cells[1].Value = newCenterPointX3;
                        dataGridView2.Rows[2].Cells[2].Value = newCenterPointY3;

                        repeat = false;
                    }
                    else
                    {
                        repeat = true;
                        counter = counter + 1;
                    }
                    centerX1 = newCenterPointX1;
                    centerY1 = newCenterPointY1;
                    centerX2 = newCenterPointX2;
                    centerY2 = newCenterPointY2;
                    centerX3 = newCenterPointX3;
                    centerY3 = newCenterPointY3;
                }
            }
            else
            {
                toolStripStatusLabel1.Text = "Unhandled Exception!";
            }
        }
    }
}
