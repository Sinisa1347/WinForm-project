using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;


namespace QuestTask
{
    public partial class RegistrationForm : Form
    {
        
        SqlConnection cn;
        public RegistrationForm()
        {
            InitializeComponent();
        }

        private void RegisterForm_Load(object sender, EventArgs e)
        {

            cn = new SqlConnection(@"Data Source=DESKTOP-S1VCV1L;Initial Catalog = QuestTaskDatabase;Integrated Security=True");

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            
            this.Hide();
           LoginForm login = new LoginForm();
            login.Show();

        }

        private void btnRegister_Click(object sender, EventArgs e)
        {
            if (txtConfirmPassword.Text != string.Empty || txtPassword.Text != string.Empty || txtUserName.Text != string.Empty)
            {
                if (txtPassword.Text == txtConfirmPassword.Text)
                {
                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand("AddToLoginAndRegistrationTable", cn);
                    sqlCommand.CommandType = CommandType.StoredProcedure;
                    sqlCommand.Parameters.AddWithValue("@username", txtUserName.Text.Trim());
                    sqlCommand.Parameters.AddWithValue("@password", txtUserName.Text.Trim());
                    sqlCommand.ExecuteNonQuery();
                    cn.Close();

                    MessageBox.Show("Your Account is created . Please login now.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //}
                }
                else
                {
                    MessageBox.Show("Please enter the same password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Please enter value in all field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
