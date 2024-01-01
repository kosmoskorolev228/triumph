using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
namespace triumph
{
    public partial class Workform : Form
    {
        public Workform()
        {
            InitializeComponent();
        }
        public SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-C3SQRDK\SQLEXPRESS;Initial Catalog=triumph;Integrated Security=True");
        SqlCommand com = new SqlCommand();
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Добро пожаловать директор, Хорошего дня");
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case 0://Заявление о приеме на работу
                    {
                        zayavlenienawork zayavlenienawork = new zayavlenienawork();
                        zayavlenienawork.Show();
                        this.Hide();
                        break;
                    }
                case 1:// Увольнение сотрудников по собственному желанию
                    {
                        yvoln yvoln = new yvoln();
                        yvoln.Show();
                        this.Hide();
                        break;
                    }
                case 2://основной отпуск
                    {
                        osnovnoy_otpysk osnovnoy_Otpysk = new osnovnoy_otpysk();
                        osnovnoy_Otpysk.Show();
                        this.Hide();
                       
                        break;
                    }
                case 3://отпуск по семейным обстоятельствам
                    {
                        otpyskposemein otpyskposemein = new otpyskposemein();
                        otpyskposemein.Show();
                        this.Hide();
                        
                        break;
                    }
                case 4://в связи с рождением ребенка
                    {
                        rebenok rebenok = new rebenok();
                        rebenok.Show();
                        this.Hide();
                        break;
                    }
                case 5://с регистрацией брака
                    {
                        brak brak = new brak();
                        brak.Show();
                        this.Hide();
                        break;
                    }
                case 6://особая категория
                    {
                        osoba osoba = new osoba();
                        osoba.Show();
                        this.Hide();
                        break;
                    }
                case 7://беременность и роды
                    {
                        rod rod = new rod();
                        rod.Show();
                        this.Hide();
                        break;
                    }
                case 8://по уходу за ребенком
                    {
                        yxodzareb yxodzareb = new yxodzareb();
                        yxodzareb.Show();
                        this.Hide();
                        break;
                    }
                case 9://по усыновлению ребенка
                    {
                        ysnovlen ysnovlen = new ysnovlen();
                        ysnovlen.Show();
                        this.Hide();
                        break;
                    }
                case 10://учебный отпуск
                    {
                        ychot ychot = new ychot();
                        ychot.Show();
                        this.Hide();
                        break;
                    }
                
                case 11://Увольнение сотрудников с выходом на пенсию
                    {
                        pesiya pesiya = new pesiya();
                        pesiya.Show();
                        this.Hide();
                        break;
                    }
                case 12://Заявление о приеме на работу в порядке перевода
                    {
                        perevod perevod = new perevod();
                        perevod.Show();
                        this.Hide();
                        break;
                    }
                case 13://Оформление трудового договора
                    {
                        tryd tryd = new tryd();
                        tryd.Show();
                        this.Hide();
                        break;
                    }
                case 14://расторжение договора
                    {
                        rastor rastor = new rastor();
                        rastor.Show();
                        this.Hide();
                        break;
                    }
            }
        }

        private void ТаблицаСотрудниковToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlCommand com = new SqlCommand("select emplo.idemplo as'Код',emplo.sname as'Фамилия',emplo.name as'Имя',emplo.pname as'Отчество',adr as'Адрес',emplo.dater as'Дата рождения',emplo.telef as'Телефон',emplo.email as'E-mail',position as'Должность' from emplo", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
            label2.Text = "Таблица сотрудников";
        }

        private void ТаблицаКлиентовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlCommand com = new SqlCommand("select idcl as'Код',sname as'Фамилия',name as'Имя',pname as'Отчество',telef as'Телефон',email as'E-mail',nameorg as 'Название организации' from client", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
            label2.Text = "Таблица клиентов";
        }

        private void ТаблицаУслугToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlCommand com = new SqlCommand("select idwork as'Код', type_of_work as'Услуги',price as'Цена(руб.)' from work", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
            label2.Text = "Перечень услуг";
        }

        private void ТабоицаЗаказовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlCommand com = new SqlCommand("select idzakaz as'Код заказа',idwork as'Код работы',area as'Количество',datezak as'Дата заказа',idcl as'Код клиента' from zakaz", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
            label2.Text = "Таблица заказов";
        }

        private void ОтчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlCommand com = new SqlCommand("select iddoc as'Код договора',idzak as'Код заказа',idwork as'Код работы',idcl as'Код клиента',price as'Цена',dopinfo as'Дополнительная информация' from document", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
            label2.Text = "Отчет";
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox1.Text.ToLower()))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void Label2_Click(object sender, EventArgs e)
        {
            if(label2.Text == "Таблица сотрудников")
            {
                cotry cotry = new cotry();
                cotry.Show();
                this.Hide();
            }
            else if (label2.Text == "Таблица клиентов")
            {
                client client = new client();
                client.Show();
                this.Hide();
            }
            else if (label2.Text == "Перечень услуг")
            {
                yslug yslug = new yslug();
                yslug.Show();
                this.Hide();
            }
            else if (label2.Text == "Таблица заказов")
            {
                zak zak = new zak();
                zak.Show();
                this.Hide();
            }
            else if (label2.Text == "Отчет")
            {
                otchet otchet = new otchet();
                otchet.Show();
                this.Hide();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
        }
    }
}
