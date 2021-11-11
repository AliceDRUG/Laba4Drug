using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laba4Drug
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int m;
            m = Convert.ToInt32(textBox1.Text);
            dataGridView1.ColumnCount = 1;
            dataGridView1.RowCount = m;
            dataGridView1.Columns.Add("1", "1");
            Random rnd = new Random();
            for (int j = 0; j < m; j++)
            {
                dataGridView1.Rows.Add();
                dataGridView1[1, j].Value = rnd.Next(1, 10000);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
