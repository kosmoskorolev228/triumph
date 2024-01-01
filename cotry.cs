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
namespace triumph
{
    public partial class cotry : Form
    {
        public cotry()
        {
            InitializeComponent();
            cot();
        }
        public SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-C3SQRDK\SQLEXPRESS;Initial Catalog=triumph;Integrated Security=True");
        SqlCommand com = new SqlCommand();
        public void cot()
        {
            SqlCommand com = new SqlCommand("select emplo.idemplo as'Код',emplo.sname as'Фамилия',emplo.name as'Имя',emplo.pname as'Отчество',adr as'Адрес',emplo.dater as'Дата рождения',emplo.telef as'Телефон',emplo.email as'E-mail',position as'Должность' from emplo", conn);
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
                string insert = "insert into emplo values (@sname,@name,@pname,@adr,@dater,@telef,@email,@position)";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand command = new SqlCommand(insert, conn))
                {
                    command.Parameters.AddWithValue("@sname", textBox2.Text);
                    command.Parameters.AddWithValue("@name", textBox3.Text);
                    command.Parameters.AddWithValue("@pname", textBox4.Text);
                    command.Parameters.AddWithValue("@adr", textBox5.Text);
                    command.Parameters.AddWithValue("@dater", textBox6.Text);
                    command.Parameters.AddWithValue("@telef", textBox7.Text);
                    command.Parameters.AddWithValue("@email", textBox8.Text);
                    command.Parameters.AddWithValue("@position", textBox9.Text);
                    command.ExecuteNonQuery();
                    cot();
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
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                string update = "update emplo set sname=@sname,name=@name,pname=@pname,adr=@adr,dater=@dater,telef=@telef,email=@email,position=@position where idemplo=@idemplo";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand com = new SqlCommand(update, conn))
                {
                    int index = dataGridView1.CurrentRow.Index;
                    com.CommandText = (update);
                    com.Parameters.AddWithValue("@idemplo", dataGridView1.Rows[index].Cells[0].Value);
                    com.Parameters.AddWithValue("@sname", dataGridView1.Rows[index].Cells[1].Value);
                    com.Parameters.AddWithValue("@name", dataGridView1.Rows[index].Cells[2].Value);
                    com.Parameters.AddWithValue("@pname", dataGridView1.Rows[index].Cells[3].Value);
                    com.Parameters.AddWithValue("@adr", dataGridView1.Rows[index].Cells[4].Value);
                    com.Parameters.AddWithValue("@dater", dataGridView1.Rows[index].Cells[5].Value);
                    com.Parameters.AddWithValue("@telef", dataGridView1.Rows[index].Cells[6].Value);
                    com.Parameters.AddWithValue("@email", dataGridView1.Rows[index].Cells[7].Value);
                    com.Parameters.AddWithValue("@position", dataGridView1.Rows[index].Cells[8].Value);
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                    conn.Close();
                    cot();
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
                string del = ("delete from emplo where idemplo='" + value + "'");
                com.Connection = conn;
                com.CommandText = (del);
                conn.Open();
                com.ExecuteNonQuery();
                conn.Close();
                com.Parameters.Clear();
                cot();
            }
            catch
            {
                MessageBox.Show("Данную запись невозможно удалить");
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Workform Workform = new Workform();
            Workform.Show();
            this.Hide();
        }
    }
}
