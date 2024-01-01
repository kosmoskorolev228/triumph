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
    public partial class zak : Form
    {
        public zak()
        {
            InitializeComponent();
            zakaz();
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

        public void zakaz()
        {
            SqlCommand com = new SqlCommand("select idzakaz as'Код заказа',idwork as'Код работы',area as'Количество',datezak as'Дата заказа',idcl as'Код клиента' from zakaz", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
            
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string insert = "insert into zakaz values (@idwork,@area,@datezak,@idcl)";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand command = new SqlCommand(insert, conn))
                {
                    command.Parameters.AddWithValue("@idwork", textBox2.Text);
                    command.Parameters.AddWithValue("@area", textBox3.Text);
                    command.Parameters.AddWithValue("@datezak", textBox4.Text);
                    command.Parameters.AddWithValue("@idcl", textBox5.Text);
                   
                    
                    command.ExecuteNonQuery();
                    zakaz();
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
          
            
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                string update = "update zakaz set idwork=@idwork,area=@area,datezak=@datezak,idcl=@idcl where idzakaz=@idzakaz";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand com = new SqlCommand(update, conn))
                {
                    int index = dataGridView1.CurrentRow.Index;
                    com.CommandText = (update);
                    com.Parameters.AddWithValue("@idzakaz", dataGridView1.Rows[index].Cells[0].Value);
                    com.Parameters.AddWithValue("@idwork", dataGridView1.Rows[index].Cells[1].Value);
                    com.Parameters.AddWithValue("@area", dataGridView1.Rows[index].Cells[2].Value);
                    com.Parameters.AddWithValue("@datezak", dataGridView1.Rows[index].Cells[3].Value);
                    com.Parameters.AddWithValue("@idcl", dataGridView1.Rows[index].Cells[4].Value);
                    
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                    conn.Close();
                    zakaz();
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
                string del = ("delete from zakaz where idzakaz='" + value + "'");
                com.Connection = conn;
                com.CommandText = (del);
                conn.Open();
                com.ExecuteNonQuery();
                conn.Close();
                com.Parameters.Clear();
                zakaz();
            }
            catch
            {
                MessageBox.Show("Данную запись невозможно удалить");
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
