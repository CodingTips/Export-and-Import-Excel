using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Read_And_Write_Exce_File
{
    public partial class Form1 : Form
    {
        WorkWithExcel oWorkWithExcel = new WorkWithExcel();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DirectoryInfo Directory = new DirectoryInfo("C:\\");
         
            oWorkWithExcel.CreateExcelFile(Directory);
            MessageBox.Show("File is created! Check drive C:\\");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = oWorkWithExcel.ReadExcel().Tables[0];
        }
    }
}
