using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace triumph
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Workform workform = new Workform();
            if (textBox1.Text == "admin" && textBox2.Text == "admin")
            {
                MessageBox.Show("Добро пожаловать администратор, Хорошего дня");

                workform.Show();
                this.Hide();

            }
            else if (textBox1.Text == "meneger" && textBox2.Text == "meneger")
            {

                //workform.справочникToolStripMenuItem.Enabled = false;
                MessageBox.Show("Добро пожаловать менеджер, Хорошего дня");

                workform.Show();
                this.Hide();

            }
            else if (textBox1.Text == "director" && textBox2.Text == "director")
            {

                MessageBox.Show("Добро пожаловать директор, Хорошего дня");

                workform.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Неправильно введен логин или пароль");
                textBox1.Clear();
                textBox2.Clear();
            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.UseSystemPasswordChar = checkBox1.Checked;
        }
    }
}
