using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace QuestTask
{
    public partial class DefaultValues : Form
    {

        public DefaultValues()
        {
            InitializeComponent();
        }

        public static string SetValueForName = "";
        public static string SetValueForLastName = "";
        public static string SetValueForAge = "";
        public static string SetValueForGender = "";
        public static string SetValueForEmail = "";
        public static string SetValueForJobType = "";
        public static string SetValueForAddress = "";

        private void btnConfirmDefaultValues_Click(object sender, EventArgs e)
        {
            
            string path = Environment.CurrentDirectory + @"\log.txt";
            System.IO.File.WriteAllText(path, string.Empty);
            FileStream f = new FileStream(path, FileMode.OpenOrCreate);
            StreamWriter s = new StreamWriter(f);

            

            s.WriteLine($"");

            SetValueForName = txtName.Text;
            s.WriteLine($"{SetValueForName}");

            SetValueForLastName = txtLastName.Text;
            s.WriteLine($"{SetValueForLastName}");

            SetValueForAge = txtAge.Text;
            s.WriteLine($"{SetValueForAge}");

            if (rbtnFemale.Checked)
            {
                SetValueForGender = "Female";
                s.WriteLine($"{SetValueForGender}");
                
            }
            else if (rbtnMale.Checked)
            {
                SetValueForGender = "Male";
                s.WriteLine($"{SetValueForGender}");
            }

            SetValueForEmail = txtEmail.Text;
            s.WriteLine($"{SetValueForEmail}");

            SetValueForJobType = txtJobType.Text;
            s.WriteLine($"{SetValueForJobType}");

            SetValueForAddress = txtAddress.Text;
            s.WriteLine($"{SetValueForAddress}");

            //closing stream writer
            s.Close();
            f.Close();
        }

        private void btnGoBack_Click(object sender, EventArgs e)
        {
            EmployeeForm registrationForm = new EmployeeForm();
            registrationForm.Show();
            this.Hide();
        }

        private void DefaultValues_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void DefaultValues_Load(object sender, EventArgs e)
        {
            if(File.Exists(Environment.CurrentDirectory + @"\log.txt"))
            {

            }
            else
            {
                File.Create(Environment.CurrentDirectory + @"\log.txt");
            }
            
            EmployeeForm employeeForm = new EmployeeForm();
            employeeForm.DefaultValuesReader();

        }
    }
}
