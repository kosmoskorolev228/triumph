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
    public partial class client : Form
    {
        public client()
        {
            InitializeComponent();
            clie();
        }
        public SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-C3SQRDK\SQLEXPRESS;Initial Catalog=triumph;Integrated Security=True");
        SqlCommand com = new SqlCommand();
        public void clie()
        {
            SqlCommand com = new SqlCommand("select idcl as'Код',sname as'Фамилия',name as'Имя',pname as'Отчество',telef as'Телефон',email as'E-mail',nameorg as 'Название организации' from client", conn);
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
                string insert = "insert into client values (@sname,@name,@pname,@telef,@email,@nameorg)";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand command = new SqlCommand(insert, conn))
                {
                    command.Parameters.AddWithValue("@sname", textBox2.Text);
                    command.Parameters.AddWithValue("@name", textBox3.Text);
                    command.Parameters.AddWithValue("@pname", textBox4.Text);
                    command.Parameters.AddWithValue("@telef", textBox5.Text);
                    command.Parameters.AddWithValue("@email", textBox6.Text);
                    command.Parameters.AddWithValue("@nameorg", textBox7.Text);
                    
                    command.ExecuteNonQuery();
                    clie();
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
            

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                string update = "update client set sname=@sname,name=@name,pname=@pname,telef=@telef,email=@email,nameorg=@nameorg where idcl=@idcl";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand com = new SqlCommand(update, conn))
                {
                    int index = dataGridView1.CurrentRow.Index;
                    com.CommandText = (update);
                    com.Parameters.AddWithValue("@idcl", dataGridView1.Rows[index].Cells[0].Value);
                    com.Parameters.AddWithValue("@sname", dataGridView1.Rows[index].Cells[1].Value);
                    com.Parameters.AddWithValue("@name", dataGridView1.Rows[index].Cells[2].Value);
                    com.Parameters.AddWithValue("@pname", dataGridView1.Rows[index].Cells[3].Value);
                    com.Parameters.AddWithValue("@telef", dataGridView1.Rows[index].Cells[4].Value);
                    com.Parameters.AddWithValue("@email", dataGridView1.Rows[index].Cells[5].Value);
                    com.Parameters.AddWithValue("@nameorg", dataGridView1.Rows[index].Cells[6].Value);
                    
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                    conn.Close();
                    clie();
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
                string del = ("delete from client where idcl='" + value + "'");
                com.Connection = conn;
                com.CommandText = (del);
                conn.Open();
                com.ExecuteNonQuery();
                conn.Close();
                com.Parameters.Clear();
                clie();
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
