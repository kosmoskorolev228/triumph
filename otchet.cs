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
    public partial class otchet : Form
    {
        public otchet()
        {
            InitializeComponent();
            otche();
            dataGridView2.Visible = false;
            button5.Visible = false;
        }
        public SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-C3SQRDK\SQLEXPRESS;Initial Catalog=triumph;Integrated Security=True");
        SqlCommand com = new SqlCommand();
        public void otche()
        {
            SqlCommand com = new SqlCommand("select iddoc as'Код договора',idzak as'Код заказа',idwork as'Код работы',idcl as'Код клиента',price as'Цена',dopinfo as'Дополнительная информация' from document", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
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

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string insert = "insert into document values (@idzak,@idwork,@idcl,@price,@dopinfo)";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand command = new SqlCommand(insert, conn))
                {
                    command.Parameters.AddWithValue("@idzak", textBox2.Text);
                    command.Parameters.AddWithValue("@idwork", textBox3.Text);
                    command.Parameters.AddWithValue("@idcl", textBox4.Text);
                    command.Parameters.AddWithValue("@price", textBox5.Text);
                    command.Parameters.AddWithValue("@dopinfo", textBox6.Text);
                    
                    command.ExecuteNonQuery();
                    otche();
                }
            }
            catch
            {
                MessageBox.Show("Проверьте заполнение текстовых полей");
            }
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                string update = "update document set idzak=@idzak,idwork=@idwork,idcl=@idcl,price=@price,dopinfo=@dopinfo where iddoc=@iddoc";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand com = new SqlCommand(update, conn))
                {
                    int index = dataGridView1.CurrentRow.Index;
                    com.CommandText = (update);
                    com.Parameters.AddWithValue("@iddoc", dataGridView1.Rows[index].Cells[0].Value);
                    com.Parameters.AddWithValue("@idzak", dataGridView1.Rows[index].Cells[1].Value);
                    com.Parameters.AddWithValue("@idwork", dataGridView1.Rows[index].Cells[2].Value);
                    com.Parameters.AddWithValue("@idcl", dataGridView1.Rows[index].Cells[3].Value);
                    com.Parameters.AddWithValue("@price", dataGridView1.Rows[index].Cells[4].Value);
                    com.Parameters.AddWithValue("@dopinfo", dataGridView1.Rows[index].Cells[5].Value);
                    
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                    conn.Close();
                    otche();
                }
            }
            catch
            {
                MessageBox.Show("Данные невозможно изменить, проверьте еще раз");
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView1.CurrentRow.Index;
                string value = Convert.ToString(dataGridView1.Rows[index].Cells[0].Value);
                string del = ("delete from document where iddoc='" + value + "'");
                com.Connection = conn;
                com.CommandText = (del);
                conn.Open();
                com.ExecuteNonQuery();
                conn.Close();
                com.Parameters.Clear();
                otche();
            }
            catch
            {
                MessageBox.Show("Данную запись невозможно удалить");
            }
        }

        private void Label2_Click(object sender, EventArgs e)
        {
            SqlCommand com = new SqlCommand("select document.iddoc as'Код договора',zakaz.idzakaz as'Код заказа',work.type_of_work as'Вид услуги',client.sname as'Фамилия клиента',client.name as'Имя клиента',client.pname as'Отчество клиента',client.email as'E-mail',client.telef as'Телефон',client.nameorg as'Наименование организации',work.price as'Цена',document.dopinfo as'Дополнительная информация' from document, zakaz, work, client where document.idzak = zakaz.idzakaz and document.idwork = work.idwork and document.idcl = client.idcl", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView2.DataSource = tbl;

            label1.Visible = false;
            label2.Visible = false;
            textBox1.Visible = false;
            //button5.Visible = true;
            //button4.Visible = false;
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false;
            textBox4.Visible = false;
            textBox5.Visible = false;
            textBox6.Visible = false;
            label3.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            label6.Visible = false;
            label7.Visible = false;
            
            //.Visible = false;
            //textBox8.Visible = false;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            button5.Visible = true;
        }
        private readonly string examp = @"D:\Работа\Крылова А.А. дипломная работа\ШаблонОтчета.docx";
        private void ReplaceWord(string stub, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stub, ReplaceWith: text);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            var cl1 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[0].Value);
            var cl2 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[1].Value);
            var cl3 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[2].Value);
            var cl4 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[3].Value);
            var cl5 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[4].Value);
            var cl6 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[5].Value);
            var cl7 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[6].Value);
            var cl8 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[7].Value);
            var cl9 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[8].Value);
            var cl10 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[9].Value);
            var cl11 = Convert.ToString(this.dataGridView2.CurrentRow.Cells[10].Value);
            //var cl12 = Convert.ToString(dataGridView2.Rows[0].Cells[11].Value);

            var wordApp = new Word.Application();
            wordApp.Visible = false;
            try
            {
                var WordDocument = wordApp.Documents.Open(examp);
                ReplaceWord("{cl1}", cl1, WordDocument);
                ReplaceWord("{cl2}", cl2, WordDocument);
                ReplaceWord("{cl3}", cl3, WordDocument);
                ReplaceWord("{cl4}", cl4, WordDocument);
                ReplaceWord("{cl5}", cl5, WordDocument);
                ReplaceWord("{cl6}", cl6, WordDocument);
                ReplaceWord("{cl7}", cl7, WordDocument);
                ReplaceWord("{cl8}", cl8, WordDocument);
                ReplaceWord("{cl9}", cl9, WordDocument);
                ReplaceWord("{cl10}", cl10, WordDocument);
                ReplaceWord("{cl11}", cl11, WordDocument);
                //ReplaceWord("{cl12}", cl12, WordDocument);

                WordDocument.SaveAs(@"D:\Работа\Крылова А.А. дипломная работа\Отчет.docx");
                wordApp.Visible = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Workform workform = new Workform();
            workform.Show();
            this.Hide();
        }
    }
}
