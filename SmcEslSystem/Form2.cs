using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace SmcEslSystem
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
           

         //   button1.DialogResult = System.Windows.Forms.DialogResult.OK;//設定button1為OK
       //     button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;//設定button為Cancel
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "smartchip")
            {
                Form1 mainPage = new Form1();

                mainPage.Show();
                this.Hide();

            }
            else
            {
                label2.Text = "密碼錯誤!!";
                label2.Visible = true;
               // textBox1.Text = "";
            }
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
         //   Form1 ddd = new Form1();
          //  ddd.Show();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            label2.Visible = false;
        }
    }
}
