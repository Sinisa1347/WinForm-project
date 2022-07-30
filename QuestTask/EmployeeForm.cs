using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Runtime.InteropServices;

namespace QuestTask
{
    public partial class EmployeeForm : System.Windows.Forms.Form
    {


        readonly string ConnectionString= @"Data Source=DESKTOP-S1VCV1L;Initial Catalog = QuestTaskDatabase;Integrated Security=True";
        
        SqlConnection sqlCon;
        SqlCommand cmd;
        SqlDataAdapter sqlDataAdapter;
        System.Data.DataTable dt;

        //int iRow;


        public EmployeeForm()
        {
            InitializeComponent();
            sqlCon = new SqlConnection(ConnectionString);
            DefaultValuesReader();
            Display();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string gender;
            if (rbtnMale.Checked)
            {
                gender = "Male";
            }
            else
            {
                gender = "Female";
            }



            if (txtName.Text == "" || txtLastName.Text == "" || txtAge.Text == "" || gender == ""|| txtEmail.Text == "" || txtJobType.Text == "" || txtAddress.Text == "")
            {
                MessageBox.Show("Please fill the blanks");
            }
            else
            {

                try
                {

                    sqlCon.Open();
                    cmd = new SqlCommand("insert into Employee (EmployeeName,EmployeeLastName,EmployeeAge,EmployeeGender,EmployeeEmail,EmployeeJobType,EmployeeAddress) " +
                        "values ('" + txtName.Text + "','" + txtLastName.Text + "','" + txtAge.Text + "','" +gender+ "','" + txtEmail.Text + "','" + txtJobType.Text + "','" + txtAddress.Text + "')", sqlCon);
                    cmd.ExecuteNonQuery();
                    sqlCon.Close();
                    MessageBox.Show("You data has been save in the Database");
                    Clear();
                    Display();
                    sqlCon.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        void Clear()
        {
            txtEmployeeId.Text="";
            txtName.Text = "";
            txtLastName.Text = "";
            txtAge.Text = "";
            rbtnMale.Checked = false;
            rbtnFemale.Checked = false;
            txtEmail.Text = "";
            txtJobType.Text = "";
            txtAddress.Text = "";
        }

        void Display()
        {
            try
            {
                sqlCon.Open();
                dt = new System.Data.DataTable();
                sqlDataAdapter = new SqlDataAdapter("select * from Employee", sqlCon);
                sqlDataAdapter.Fill(dt);
                DatabaseGridView.DataSource = dt;
                sqlCon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DatabaseGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)//ceo tekst ispod govori za jedan red DataGridTable-a
        {
            DataGridViewDoubleClick(sender,e);
        }
        private void OpenedExcelDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewDoubleClick(sender,e);
        }

        public void DataGridViewDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            txtEmployeeId.Text = DatabaseGridView.Rows[e.RowIndex].Cells[0].Value.ToString();
            txtName.Text = DatabaseGridView.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtLastName.Text = DatabaseGridView.Rows[e.RowIndex].Cells[2].Value.ToString();
            txtAge.Text = DatabaseGridView.Rows[e.RowIndex].Cells[3].Value.ToString();
            rbtnMale.Checked = true;
            rbtnFemale.Checked = false;


            if (DatabaseGridView.Rows[e.RowIndex].Cells[4].Value.ToString() == "Female")
            {
                rbtnMale.Checked = false;
                rbtnFemale.Checked = true;
            }
            txtEmail.Text = DatabaseGridView.Rows[e.RowIndex].Cells[5].Value.ToString();
            txtJobType.Text = DatabaseGridView.Rows[e.RowIndex].Cells[6].Value.ToString();
            txtAddress.Text = DatabaseGridView.Rows[e.RowIndex].Cells[7].Value.ToString();
            btnUpdate.Enabled = true;
            btnDelete.Enabled = true;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                string gender;
                if (rbtnMale.Checked)
                {
                    gender = "Male";
                }
                else
                {
                    gender = "Female";
                }

                sqlCon.Open();
                cmd = new SqlCommand("update Employee set EmployeeName='" + txtName.Text + "',EmployeeLastName='" + txtLastName.Text + "',EmployeeAge='" + txtAge.Text + "',EmployeeGender='" + gender+ "',EmployeeEmail='" + txtEmail.Text + "',EmployeeJobType='" + txtJobType.Text + "',EmployeeAddress='" + txtAddress.Text + "' where EmployeeId='" + txtEmployeeId.Text + "'", sqlCon);
                cmd.ExecuteNonQuery();
                sqlCon.Close();
                MessageBox.Show("You data has been updated");
                Display();
                Clear();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                sqlCon.Open();
                cmd = new SqlCommand("delete from Employee where EmployeeId='" + txtEmployeeId.Text + "'", sqlCon);
                cmd.ExecuteNonQuery();
                sqlCon.Close();
                
                MessageBox.Show("Your record has been deleted");
                Display();
                Clear();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excell = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb = Excell.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet ws = (Worksheet)Excell.ActiveSheet;
                Excell.Visible = true;
                int rowCount = DatabaseGridView.Rows.Count;


                for (int j = 2; j < DatabaseGridView.Rows.Count; j++)
                {
                    for (int i = 1; i <= 1; i++)
                    {
                        ws.Cells[j, i] = DatabaseGridView.Rows[j - 2].Cells[i - 1].Value;
                    }
                }

                for (int i = 1; i < DatabaseGridView.Columns.Count + 1; i++)
                {
                    ws.Cells[1, i] = DatabaseGridView.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < DatabaseGridView.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < DatabaseGridView.Columns.Count; j++)
                    {
                        ws.Cells[i + 2, j + 1] = DatabaseGridView.Rows[i].Cells[j].Value;
                    }
                }
            }
            catch (Exception)
            {

            }
        }

        string sFileName;

        private void btnClearData_Click(object sender, EventArgs e)
        {
            Clear();
        }

        private void btnDefaultValues_Click(object sender, EventArgs e)
        {
            this.Hide();
            DefaultValues defaultValues = new DefaultValues();
            defaultValues.Show();
            
        }

        private void RegistrationForm_Load(object sender, EventArgs e)
        {
            LoadRecentList();
            LoadNumberOfRecentProjects();
            DefaultValuesReader();
            Display();

            foreach (string item in MRUlist)
            {
                ToolStripMenuItem fileRecent= new ToolStripMenuItem(item,null,RecentFile_click);
                RecentToolStripMenuItem.DropDownItems.Add(fileRecent);
            }
        }



        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        private void readExcel(string sFile)
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sFile);           // WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets["Sheet1"];      // NAME OF THE SHEET.

            string file = sFile; //variable for the Excel File Location
            System.Data.DataTable dt = new System.Data.DataTable(); //container for our excel data
            DataRow row;

            //file = openFileDialog1.FileName; //get the filename with the location of the file

            //Create Object for Microsoft.Office.Interop.Excel that will be use to read excel file

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file);

            Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];

            Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

            int rowCount = excelRange.Rows.Count; //get row count of excel data

            int colCount = excelRange.Columns.Count; // get column count of excel data

            //Get the first Column of excel file which is the Column Name

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                }
                break;
            }

            //Get Row Data of Excel

            int rowCounter; //This variable is used for row index number
            for (int i = 2; i <= rowCount; i++) //Loop for available row of excel data
            {
                row = dt.NewRow(); //assign new row to DataTable
                rowCounter = 0;
                for (int j = 1; j <= colCount; j++) //Loop for available column of excel data
                {
                    //check if cell is empty
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    {
                        row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                    }
                    else
                    {
                        row[i] = "";
                    }
                    rowCounter++;
                }
                dt.Rows.Add(row); //add row to DataTable
            }

            OpenedExcelDataGridView.DataSource = dt; //assign DataTable as Datasource for DataGridview

            //close and clean excel process
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(excelWorksheet);
            //quit apps
            excelWorkbook.Close();
            Marshal.ReleaseComObject(excelWorkbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);


            //int iRow;

            //for (iRow = 2; iRow <= xlWorkSheet.Rows.Count; iRow++)  // START FROM THE SECOND ROW.
            //{
            //    if (xlWorkSheet.Cells[iRow, 1].Text == null)
            //    {
            //        break;// BREAK LOOP.
            //    }
            //    else
            //    {   // POPULATE COMBO BOX.
            //        txtEmployeeId.Text = xlWorkSheet.Cells[iRow, 1].Text;
            //        txtName.Text=xlWorkSheet.Cells[iRow, 2].Text;
            //        txtLastName.Text = xlWorkSheet.Cells[iRow, 3].Text;
            //        txtAge.Text = xlWorkSheet.Cells[iRow, 4].Text;

            //        txtEmail.Text = xlWorkSheet.Cells[iRow, 6].Text;
            //        txtJobType.Text = xlWorkSheet.Cells[iRow, 7].Text;
            //        txtAddress.Text = xlWorkSheet.Cells[iRow, 8].Text;
            //        if (xlWorkSheet.Cells[iRow, 5].Value== "Male")
            //        {
            //            rbtnMale.Checked = true;
            //            rbtnFemale.Checked = false;
            //            //break;
            //        }
            //        else
            //        {
            //            rbtnFemale.Checked = true;
            //            rbtnMale.Checked = false;
            //            //break;
            //        }
            //    }
            //}
            xlWorkBook.Close();
                xlApp.Quit();

        }

        void btnNewProject_Click(object sender, EventArgs e)
        {
            if(txtName!=null)
            {
                Clear();
                DefaultValuesReader();
            }
            else
            {
                DefaultValuesReader();
            }
            
        }

        public void DefaultValuesReader()
        {
            if (File.Exists(Environment.CurrentDirectory + @"\log.txt"))
            {
                using (StreamReader reader = new StreamReader(new FileStream(Environment.CurrentDirectory + @"\log.txt", FileMode.Open)))
                {
                    string line;
                    int counter = 0;
                    // Read line by line  
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (counter == 0)
                        {
                            counter++;
                        }
                        else if (counter == 1)
                        {
                            txtName.Text = line;
                            counter++;
                        }

                        else if (counter == 2)
                        {
                            txtLastName.Text = line;
                            counter++;
                        }
                        else if (counter == 3)
                        {
                            txtAge.Text = line;
                            counter++;
                        }
                        else if (counter == 4)
                        {
                            if (line == "Male")
                            {
                                rbtnMale.Checked = true;
                                rbtnFemale.Checked = false;
                                counter++;

                            }
                            else if (line == "Female")
                            {
                                rbtnMale.Checked = false;
                                rbtnFemale.Checked = true;
                                counter++;

                            }
                            else
                            {
                                counter++;
                            }

                        }
                        else if (counter == 5)
                        {
                            txtEmail.Text = line;
                            counter++;
                        }
                        else if (counter == 6)
                        {
                            txtJobType.Text = line;
                            counter++;
                        }
                        else if (counter == 7)
                        {
                            txtAddress.Text = line;
                            break;
                        }
                    }
                }
            }
            else
            {
                Clear();
            }
        }
        

        private readonly Queue<string> MRUlist = new Queue<string>();

        private void SaveRecentFile(string strPath)
        {
            RecentToolStripMenuItem.DropDownItems.Clear();

            LoadRecentList();

            if (!(MRUlist.Contains(strPath)))

                MRUlist.Enqueue(strPath);

            if (txtNumberOfRecentFiles == null)
            {
                int counter = 5;

                while (MRUlist.Count > counter)

                    MRUlist.Dequeue();

                foreach (string strItem in MRUlist)
                {
                    ToolStripMenuItem tsRecent = new
                       ToolStripMenuItem(strItem, null);

                    RecentToolStripMenuItem.DropDownItems.Add(tsRecent);
                }

                StreamWriter stringToWrite = new
                   StreamWriter(System.Environment.CurrentDirectory +
                   @"\Recent.txt");

                foreach (string item in MRUlist)

                    stringToWrite.WriteLine(item);

                stringToWrite.Flush();

                stringToWrite.Close();
            }
            else
            {
                int counter = Convert.ToInt32(txtNumberOfRecentFiles.Text);

                while (MRUlist.Count > counter)
                {
                    MRUlist.Dequeue();
                }
                foreach (string strItem in MRUlist)
                {
                    ToolStripMenuItem tsRecent = new
                       ToolStripMenuItem(strItem, null);

                    RecentToolStripMenuItem.DropDownItems.Add(tsRecent);
                }

                StreamWriter stringToWrite = new StreamWriter(System.Environment.CurrentDirectory + @"\Recent.txt");

                foreach (string item in MRUlist)

                    stringToWrite.WriteLine(item);

                stringToWrite.Flush();

                stringToWrite.Close();
            }
        }
        private void LoadRecentList()
        {
            MRUlist.Clear();
            try
            {
                StreamReader srStream = new StreamReader
                   (Environment.CurrentDirectory + @"\Recent.txt");

                string strLine = "";

                while ((InlineAssignHelper(ref strLine,
                      srStream.ReadLine())) != null)

                    MRUlist.Enqueue(strLine);

                srStream.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Creating Recent.txt in Debug folder");
                File.Create(Environment.CurrentDirectory + @"\Recent.txt");
            }
        }
        private void RecentFile_click(object sender, EventArgs e)
        {
            var sFileName=sender.ToString();
            if (sFileName.Trim() != "")
            {
                readExcel(sFileName);

            }
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)//open excel files
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Excel File to Edit";
            openFileDialog.FileName = "";
            openFileDialog.Filter = "Excel File|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                sFileName = openFileDialog.FileName;

                if (sFileName.Trim() != "")
                {
                    readExcel(sFileName);
                    SaveRecentFile(sFileName);

                }
            }
        }
        private static T InlineAssignHelper<T>(ref T target, T value)
        {
            target = value;
            return value;
        }

        private void txtChangeNumberOfRecentFiles_Click(object sender, EventArgs e)
        {
            string numberOfRecentProjectsPath = Environment.CurrentDirectory + @"\NumberOfRecentProjects.txt";
            using (StreamWriter streamWriter = new StreamWriter(numberOfRecentProjectsPath))
            {
                if (txtNumberOfRecentFiles.Text == null)
                {

                    streamWriter.WriteLine(5);

                }
                else
                {
                    streamWriter.WriteLine(txtNumberOfRecentFiles.Text);
                }
            }
        }
        public void LoadNumberOfRecentProjects()
        {
            if (File.Exists(Environment.CurrentDirectory + @"\NumberOfRecentProjects.txt"))
            {
                using (StreamReader streamReader = new StreamReader(Environment.CurrentDirectory + @"\NumberOfRecentProjects.txt"))
                {
                    txtNumberOfRecentFiles.Text = streamReader.ReadToEnd();
                }
            }
            else
            {
                File.Create(Environment.CurrentDirectory + @"\NumberOfRecentProjects.txt");
                using (StreamReader streamReader = new StreamReader(Environment.CurrentDirectory + @"\NumberOfRecentProjects.txt"))
                {
                    txtNumberOfRecentFiles.Text = streamReader.ReadToEnd();
                }
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)//recent excel
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Excel File to Edit";
            openFileDialog.FileName = "";
            openFileDialog.Filter = "Excel File|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                sFileName = openFileDialog.FileName;

                if (sFileName.Trim() != "")
                {
                    readExcel(sFileName);

                }
            }
        }

        private void btnOpenedExcelToDatabase_Click(object sender, EventArgs e)
        {
            sqlCon.Open();
            for (int i = 0; i < OpenedExcelDataGridView.Rows.Count-1; i++)
            {
                cmd = new SqlCommand("insert into Employee (EmployeeName,EmployeeLastName,EmployeeAge,EmployeeGender,EmployeeEmail,EmployeeJobType,EmployeeAddress) " +
                "values ('" + OpenedExcelDataGridView.Rows[i].Cells["EmployeeName"].Value + "','" + OpenedExcelDataGridView.Rows[i].Cells["EmployeeLastName"].Value + "','" + OpenedExcelDataGridView.Rows[i].Cells["EmployeeAge"].Value + "','" + OpenedExcelDataGridView.Rows[i].Cells["EmployeeGender"].Value + "','" + OpenedExcelDataGridView.Rows[i].Cells["EmployeeEmail"].Value + "','" + OpenedExcelDataGridView.Rows[i].Cells["EmployeeJobType"].Value + "','" + OpenedExcelDataGridView.Rows[i].Cells["EmployeeAddress"].Value + "')", sqlCon);
                cmd.ExecuteNonQuery();
            }
            sqlCon.Close();
            Display();

        }

        //private void EmployeeForm_FormClosing(object sender, FormClosingEventArgs e)
        //{
        //    DefaultValues defaultValues = new DefaultValues();

        //    if(Form.ActiveForm==DefaultValues.ActiveForm)
        //    {

        //    }
        //    //EmployeeForm employeeForm = new EmployeeForm();
        //    else
        //    {
        //        System.Windows.Forms.Application.Exit();
        //    }

        //}


    }
}
