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

namespace QuestTask
{
    public partial class LoginForm : Form
    {
        SqlConnection cn;

        public LoginForm()
        {
            InitializeComponent();

        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
            cn = new SqlConnection(@"Data Source=DESKTOP-S1VCV1L;Initial Catalog = QuestTaskDatabase;Integrated Security=True");
            //cn.Open();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (txtPassword.Text != string.Empty || txtUserName.Text != string.Empty)
            {
                cn.Open();
                
                SqlCommand cmd = new SqlCommand("select * from LoginAndRegistration where username='" + txtUserName.Text + "' and password='" + txtPassword.Text + "'", cn);
                if (cmd.ExecuteReader().Read())
                {
                    cn.Close();
                    this.Hide();
                    EmployeeForm home = new EmployeeForm();
                    home.ShowDialog();
                }
                else
                {
                    //cmd.ExecuteReader().Close();
                    cn.Close();
                    MessageBox.Show("No Account avilable with this username and password ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("Please enter value in all field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cn.Close();
            }

        }

        private void btnRegister_Click(object sender, EventArgs e)
        {

            this.Hide();
            //cn.Close();
            RegistrationForm registration = new RegistrationForm();
            registration.ShowDialog();
        }
    }
}
