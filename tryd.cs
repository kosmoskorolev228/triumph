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
    public partial class tryd : Form
    {
        public tryd()
        {
            InitializeComponent();
            osnovnoy();
        }
        public SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-C3SQRDK\SQLEXPRESS;Initial Catalog=triumph;Integrated Security=True");
        SqlCommand com = new SqlCommand();
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
        public void osnovnoy()
        {
            SqlCommand com = new SqlCommand("select idos as'Номер трудового договора',nameorg as'Наименование организации',inn as'ИНН',bik as'БИК',kpp as'КПП',address as'Адрес',datan as'Дата начала',fiodir as'Фио директора',fiocotr as'Фио сотрудника',pasportdate as'Паспортные данные',zp as'Заработанная плата' from TD", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Workform workform = new Workform();
            workform.Show();
            this.Hide();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                string insert = "insert into TD values (@nameorg,@inn,@bik,@kpp,@address,@datan,@fiodir,@fiocotr,@pasportdate,@zp)";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand command = new SqlCommand(insert, conn))
                {
                    command.Parameters.AddWithValue("@nameorg", textBox2.Text);
                    command.Parameters.AddWithValue("@inn", textBox3.Text);
                    command.Parameters.AddWithValue("@bik", textBox4.Text);
                    command.Parameters.AddWithValue("@kpp", textBox5.Text);
                    command.Parameters.AddWithValue("@address", textBox6.Text);
                    command.Parameters.AddWithValue("@datan", textBox7.Text);
                    command.Parameters.AddWithValue("@fiodir", textBox8.Text);
                    command.Parameters.AddWithValue("@fiocotr", textBox9.Text);
                    command.Parameters.AddWithValue("@pasportdate", textBox10.Text);
                    command.Parameters.AddWithValue("@zp", textBox11.Text);
                    
                    command.ExecuteNonQuery();
                    osnovnoy();
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
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();
            textBox10.Clear();
            textBox11.Clear();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            try
            {
                string update = "update TD set nameorg=@nameorg,inn=@inn,bik=@bik,kpp=@kpp,address=@address,datan=@datan,fiodir=@fiodir, fiocotr=@fiocotr,pasportdate=@pasportdate,zp=@zp where idos=@idos";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand com = new SqlCommand(update, conn))
                {
                    int index = dataGridView1.CurrentRow.Index;
                    com.CommandText = (update);
                    com.Parameters.AddWithValue("@idos", dataGridView1.Rows[index].Cells[0].Value);
                    com.Parameters.AddWithValue("@nameorg", dataGridView1.Rows[index].Cells[1].Value);
                    com.Parameters.AddWithValue("@inn", dataGridView1.Rows[index].Cells[2].Value);
                    com.Parameters.AddWithValue("@bik", dataGridView1.Rows[index].Cells[3].Value);
                    com.Parameters.AddWithValue("@kpp", dataGridView1.Rows[index].Cells[4].Value);
                    com.Parameters.AddWithValue("@address", dataGridView1.Rows[index].Cells[5].Value);
                    com.Parameters.AddWithValue("@datan", dataGridView1.Rows[index].Cells[6].Value);
                    com.Parameters.AddWithValue("@fiodir", dataGridView1.Rows[index].Cells[7].Value);
                    com.Parameters.AddWithValue("@fiocotr", dataGridView1.Rows[index].Cells[8].Value);
                    com.Parameters.AddWithValue("@pasportdate", dataGridView1.Rows[index].Cells[9].Value);
                    com.Parameters.AddWithValue("@zp", dataGridView1.Rows[index].Cells[10].Value);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                    conn.Close();
                    osnovnoy();
                }
            }
            catch
            {
                MessageBox.Show("Данные невозможно изменить, проверьте еще раз");
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView1.CurrentRow.Index;
                string value = Convert.ToString(dataGridView1.Rows[index].Cells[0].Value);
                string del = ("delete from TD where idos='" + value + "'");
                com.Connection = conn;
                com.CommandText = (del);
                conn.Open();
                com.ExecuteNonQuery();
                conn.Close();
                com.Parameters.Clear();
                osnovnoy();
            }
            catch
            {
                MessageBox.Show("Данную запись невозможно удалить");
            }
        }
        private readonly string examp = @"D:\Работа\Крылова А.А. дипломная работа\ШаблонТрудовогоДоговора.docx";
        private void ReplaceWord(string stub, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stub, ReplaceWith: text);
        }
        private void Button5_Click(object sender, EventArgs e)
        {
            var cl1 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[0].Value);
            var cl2 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[1].Value);
            var cl3 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[7].Value);
            var cl4 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[8].Value);
            var cl5 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[1].Value);
            var cl6 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[5].Value);
            var cl7 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[7].Value);
            var cl8 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[6].Value);
            var cl9 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[10].Value);
            var cl10 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[7].Value);
            var cl11 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[5].Value);
            var cl12 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[2].Value);
            var cl13 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[3].Value);
            var cl14 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[4].Value);
            var cl15 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[7].Value);
            var cl16 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[9].Value);
            var cl17 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[7].Value);
            var cl18 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[8].Value);
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
                ReplaceWord("{cl12}", cl12, WordDocument);
                ReplaceWord("{cl13}", cl13, WordDocument);
                ReplaceWord("{cl14}", cl14, WordDocument);
                ReplaceWord("{cl15}", cl15, WordDocument);
                ReplaceWord("{cl16}", cl16, WordDocument);
                ReplaceWord("{cl17}", cl17, WordDocument);
                ReplaceWord("{cl18}", cl18, WordDocument);
                WordDocument.SaveAs(@"D:\Работа\Крылова А.А. дипломная работа\ДоговорТрудовогоДоговора.docx");
                wordApp.Visible = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }
    }
}
