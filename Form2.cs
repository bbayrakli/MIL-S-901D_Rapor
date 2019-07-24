using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _901DD
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            textBox2.PasswordChar = '*';
            textBox2.MaxLength = 10;
           
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
        string username = "burak", parola = "Kuleli5075";

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox2.PasswordChar = '\0';
            }
            else
            {
                textBox2.PasswordChar = '*';
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == username && textBox2.Text == parola)
            {
                this.Hide();
                Form Form1 = new Form1();
                Form1.ShowDialog();
               
            }
            else
            {
                MessageBox.Show("geçersiz Kullanıcı Adı ya da Parola");
                
                System.Environment.Exit(0);
                
            }
        }
    }
}
