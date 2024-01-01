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
    public partial class osnovnoy_otpysk : Form
    {
        public SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-C3SQRDK\SQLEXPRESS;Initial Catalog=triumph;Integrated Security=True");
        SqlCommand com = new SqlCommand();
        public osnovnoy_otpysk()
        {
            InitializeComponent();
            osnovnoy();
        }
        public void osnovnoy()
        {
            SqlCommand com = new SqlCommand("select idos as'Номер отпуска',nameorg as'Наименование организации',fiodir as'Фио директора',fiocot as'Фио сотрудника',position as'Должность',daten as'Дата начала',datek as'Дата окончания',kolday as'Количество дней' from osnovnoyotpysk", conn);
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
            Workform workform = new Workform();
            workform.Show();
            this.Hide();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                string insert = "insert into osnovnoyotpysk values (@nameorg,@fiodir,@fiocot,@position,@daten,@datek,@kolday)";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand command = new SqlCommand(insert, conn))
                {
                    command.Parameters.AddWithValue("@nameorg", textBox2.Text);
                    command.Parameters.AddWithValue("@fiodir", textBox3.Text);
                    command.Parameters.AddWithValue("@fiocot", textBox4.Text);
                    command.Parameters.AddWithValue("@position", textBox5.Text);
                    command.Parameters.AddWithValue("@daten", textBox6.Text);
                    command.Parameters.AddWithValue("@datek", textBox7.Text);
                    command.Parameters.AddWithValue("@kolday", textBox8.Text);
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
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView1.CurrentRow.Index;
                string value = Convert.ToString(dataGridView1.Rows[index].Cells[0].Value);
                string del = ("delete from osnovnoyotpysk where idos='" + value + "'");
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

        private void Button3_Click(object sender, EventArgs e)
        {
            try
            {
                string update = "update osnovnoyotpysk set nameorg=@nameorg,fiodir=@fiodir,fiocot=@fiocot,position=@position,daten=@daten,datek=@datek,kolday=@kolday where idos=@idos";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand com = new SqlCommand(update, conn))
                {
                    int index = dataGridView1.CurrentRow.Index;
                    com.CommandText = (update);
                    com.Parameters.AddWithValue("@idos", dataGridView1.Rows[index].Cells[0].Value);
                    com.Parameters.AddWithValue("@nameorg", dataGridView1.Rows[index].Cells[1].Value);
                    com.Parameters.AddWithValue("@fiodir", dataGridView1.Rows[index].Cells[2].Value);
                    com.Parameters.AddWithValue("@fiocot", dataGridView1.Rows[index].Cells[3].Value);
                    com.Parameters.AddWithValue("@position", dataGridView1.Rows[index].Cells[4].Value);
                    com.Parameters.AddWithValue("@daten", dataGridView1.Rows[index].Cells[5].Value);
                    com.Parameters.AddWithValue("@datek", dataGridView1.Rows[index].Cells[6].Value);
                    com.Parameters.AddWithValue("@kolday", dataGridView1.Rows[index].Cells[7].Value);
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
        private readonly string examp = @"D:\Работа\Крылова А.А. дипломная работа\ШаблонОтпускаОсновного1.docx";
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
            var cl3 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[2].Value);
            var cl4 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[3].Value);
            var cl5 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[4].Value);
            var cl6 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[5].Value);
            var cl7 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[6].Value);
            var cl8 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[7].Value);

            var cl9 = Convert.ToString(this.dataGridView1.CurrentRow.Cells[3].Value);

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

                WordDocument.SaveAs(@"D:\Работа\Крылова А.А. дипломная работа\ДоговорОтпускаОсновного.docx");
                wordApp.Visible = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }
        
    }
}
