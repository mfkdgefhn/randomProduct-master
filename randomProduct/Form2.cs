using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace randomProduct
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string name = this.textBox1.Text.Trim();
            double p =0.0;
            try
            {
                p = Convert.ToDouble(this.textBox2.Text.Trim()); 
                if (p > 1)
                {
                    MessageBox.Show("概率不得大于1");
                }
                else
                {
                    Form1.customerName = name;
                    Form1.probability = p;
                    this.Close();
                }
            }
            catch
            {
                MessageBox.Show("不是请输入0-1之前的小数");
                return;
            }
            
        }
    }
}
