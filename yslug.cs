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
    public partial class yslug : Form
    {
        public yslug()
        {
            InitializeComponent();
            ysl();
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
        public void ysl()
        {
            SqlCommand com = new SqlCommand("select idwork as'Код', type_of_work as'Услуги',price as'Цена(руб.)' from work", conn);
            SqlDataAdapter ad = new SqlDataAdapter(com);
            DataTable tbl = new DataTable();
            ad.Fill(tbl);
            dataGridView1.DataSource = tbl;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string insert = "insert into work values (@type_of_work,@price)";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand command = new SqlCommand(insert, conn))
                {
                    command.Parameters.AddWithValue("@type_of_work", textBox2.Text);
                    command.Parameters.AddWithValue("@price", textBox3.Text);
                  

                    command.ExecuteNonQuery();
                    ysl();
                }
            }
            catch
            {
                MessageBox.Show("Проверьте заполнение текстовых полей");
            }
            textBox2.Clear();
            textBox3.Clear();
            
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                string update = "update work set type_of_work=@type_of_work,price=@price where idwork=@idwork";
                com.Connection = conn;
                conn.Open();
                using (SqlCommand com = new SqlCommand(update, conn))
                {
                    int index = dataGridView1.CurrentRow.Index;
                    com.CommandText = (update);
                    com.Parameters.AddWithValue("@idwork", dataGridView1.Rows[index].Cells[0].Value);
                    com.Parameters.AddWithValue("@type_of_work", dataGridView1.Rows[index].Cells[1].Value);
                    com.Parameters.AddWithValue("@price", dataGridView1.Rows[index].Cells[2].Value);
                    
                    com.ExecuteNonQuery();
                    com.Parameters.Clear();
                    conn.Close();
                    ysl();
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
                string del = ("delete from work where idwork='" + value + "'");
                com.Connection = conn;
                com.CommandText = (del);
                conn.Open();
                com.ExecuteNonQuery();
                conn.Close();
                com.Parameters.Clear();
                ysl();
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
