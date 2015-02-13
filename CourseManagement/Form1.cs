using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace CourseManagement
{
    public partial class CourseManagement : Form
    {
        public int complete, total;
        public String appPath = System.IO.Path.GetFullPath("..\\..\\").ToString();
        public CourseManagement()
        {

            Thread t = new Thread(new ThreadStart(SplashScreen));
            t.Start();
            Thread.Sleep(5000);
            InitializeComponent();
            t.Abort();
            this.Activate();
            this.Focus();
            this.TopMost = true;
            this.Show();
            this.BringToFront();
            this.GotFocus += Form_GotFocus;
            this.Select();
            this.dataTable1TableAdapter1.Fill(this.dataSet2.DataTable1);
            this.reportViewer2.RefreshReport();

            CourseCodeTextBox.Enabled = false;
            CourseNameTextBox.Enabled = false;
            CurriculumTextBox.Enabled = false;
            CourseDesTextBox8.Enabled = false;
            CourseEquivalentTextBox.Enabled = false;
            CourseCategoryComboBox.Enabled = false;
            CoursePreTextBox.Enabled = false;
            CourseCreditComboBox.Enabled = false;

            IDTextBox.Enabled = false;
            FirstNameTextBox.Enabled = false;
            SurnameTextBox.Enabled = false;
            FacultyComboBox.Enabled = false;
            MajorComboBox.Enabled = false;
            StatusComboBox.Enabled = false;
            DegreeComboBox.Enabled = false;
            GPATextBox.Enabled = false;
            AcademicTextBox.Enabled = false;
            NationalityTextBox.Enabled = false;
            ReligionTextBox.Enabled = false;
            AddressTextBox.Enabled = false;
            CurriculumComboBox.Enabled = false;
            PhonetextBox.Enabled = false;
            EmailTextBox.Enabled = false;
            CreditStudentShowTextBox.Enabled = false;
            ((Control)this.CourseResultTap).Enabled = false;
            ModifyLecturerNameTextBox.Enabled = false;
            ModifyLecturerTypeComboBox.Enabled = false;
            ModifyCurriculumTextBox.Enabled = false;
            AddCurriculumYearTextBox.MaxLength = 6;
            ModifyCurriculumTextBox.MaxLength = 6;
            ((Control)this.CourseManageTap).Enabled = false;
            ((Control)this.Course_Detail_Tab).Enabled = false;

            int count, countLecturer;
            ArrayList curriculumYearList = new ArrayList();
            ArrayList lecturerList = new ArrayList();
            OleDbConnection conn = new OleDbConnection();

            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String[] curriculumYear;

                String query = "SELECT * FROM Curriculum";
                //String queryCount = "SELECT COUNT(*) FROM Course";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                //OleDbCommand cmd2 = new OleDbCommand(queryCount, conn);

                cmd.ExecuteNonQuery();

                OleDbDataReader reader = cmd.ExecuteReader();
                //OleDbDataReader readCount = cmd2.ExecuteReader();

                OleDbCommand command = conn.CreateCommand();
                command.CommandText = "SELECT COUNT(*) FROM Curriculum";
                count = (int)command.ExecuteScalar();
                curriculumYear = new String[count];
                while (reader.Read()) 
                {
                    curriculumYearList.Add(reader.GetString(1));
                }
                curriculumYearList.Sort();
                curriculumYear = (String[])curriculumYearList.ToArray(typeof(string));
                for (int i = 0; i < count; i++)
                {
                    CurriculumStudentComboBox.Items.Add(curriculumYear[i]);
                    CurriculumComboBox.Items.Add(curriculumYear[i]);
                    CurriculumTextBox.Items.Add(curriculumYear[i]);
                    CurriculumAddtextBox.Items.Add(curriculumYear[i]);
                }

                string query2 = "SELECT DISTINCT (Course.Course_Code), Course.Course_Name, (SELECT COUNT(*) FROM Student WHERE Student.Student_Syllabus = Course.Year_of_Syllabus OR Course.Year_of_Syllabus = 'ALL')-(SELECT COUNT(Course_Code) FROM Course_Registration WHERE Course_Registration.Course_Code = Course.Course_Code) AS Incomplete, (SELECT COUNT(Course_Code) FROM Course_Registration WHERE Course.Course_Code = Course_Registration.Course_Code AND (Course.Year_of_Syllabus = Student.Student_Syllabus OR Course.Year_of_Syllabus = 'ALL' )) AS Completed FROM Course INNER JOIN Student ON Course.Year_of_Syllabus = Student.Student_Syllabus OR Course.Year_of_Syllabus = 'All'";

                OleDbCommand cmdGrid2 = new OleDbCommand(query2, conn);
                cmdGrid2.CommandType = CommandType.Text;
                OleDbDataAdapter d2 = new OleDbDataAdapter(cmdGrid2);
                DataTable unregistered = new DataTable();
                d2.Fill(unregistered);
                CoursesSelectdataGridView3.DataSource = unregistered;
                this.CoursesSelectdataGridView3.Sort(this.CoursesSelectdataGridView3.Columns["Incomplete"], ListSortDirection.Descending);

            }
            catch (Exception ex)
            {
               MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                try
                {
                    string strSql = "SELECT Course_Offer.Course_Code, Course.Course_Name FROM Course_Offer INNER JOIN Course ON Course_Offer.Course_Code = Course.Course_Code";
                    OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                    cmdGrid.CommandType = CommandType.Text;
                    OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                    DataTable registered = new DataTable();
                    da.Fill(registered);
                    CoursesOfferListdataGridView.DataSource = registered;
                    this.CoursesOfferListdataGridView.Sort(this.CoursesOfferListdataGridView.Columns["Course_Code"], ListSortDirection.Ascending);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                finally
                {
                    try
                    {
                        string strSql = "SELECT Course_Code FROM Course_Offer";
                        OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                        cmdGrid.CommandType = CommandType.Text;
                        OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                        DataTable registered = new DataTable();
                        da.Fill(registered);
                        dataGridView2.DataSource = registered;
                        this.dataGridView2.Sort(this.dataGridView2.Columns["Course_Code"], ListSortDirection.Ascending);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                    }
                    finally
                    {
                        CoursesOfferListdataGridView.CellMouseClick += new    DataGridViewCellMouseEventHandler(CourseNameClick);
                        CoursesSelectdataGridView3.CellMouseDoubleClick += new DataGridViewCellMouseEventHandler(CourseNameClick1);
                    }
                }
                conn.Close();
            }
            try
            {
                conn.Open();
                String[] lecturer;

                String queryLecturer = "SELECT * FROM Lecturer";

                OleDbCommand cmdLecturer = new OleDbCommand(queryLecturer, conn);

                cmdLecturer.ExecuteNonQuery();

                OleDbDataReader readerLecturer = cmdLecturer.ExecuteReader();

                OleDbCommand commandLecturer = conn.CreateCommand();
                commandLecturer.CommandText = "SELECT COUNT(*) FROM Lecturer";
                countLecturer = (int)commandLecturer.ExecuteScalar();
                lecturer = new String[countLecturer];
                while (readerLecturer.Read())
                {
                    lecturerList.Add(readerLecturer.GetString(1));
                }
                lecturerList.Sort();
                lecturer = (String[])lecturerList.ToArray(typeof(string));
                for (int i = 0; i < countLecturer; i++)
                {
                    comboBox3.Items.Add(lecturer[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }
        }

        private void Form_GotFocus(object sender, EventArgs e)
        {
            if (!this.Visible)
            {
                this.Show();
            }
        }

        public void SplashScreen()
        {
            Application.Run(new CourseManagementProject.Splash_Screen());
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void Search_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dataSet2.DataTable1' table. You can move, or remove it, as needed.
            this.dataTable1TableAdapter1.Fill(this.dataSet2.DataTable1);
            // TODO: This line of code loads data into the 'dataSet1.DataTable1' table. You can move, or remove it, as needed.
            this.dataTable1TableAdapter.Fill(this.dataSet1.DataTable1);
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void StudentListTabPage_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void StudentProfileTap_Click(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void FirstNameLabel_Click(object sender, EventArgs e)
        {

        }

        private void FacultyLabel_Click(object sender, EventArgs e)
        {

        }

        private void StatusLabel_Click(object sender, EventArgs e)
        {

        }

        private void GPALabel_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void ReportsTapPage_Click(object sender, EventArgs e)
        {

        }

        private void reportViewer2_Load(object sender, EventArgs e)
        {

        }

        private void CourseInfoTap_Click(object sender, EventArgs e)
        {

        }

        private void AddCoursebutton_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String courseCode = CourseCodeAddtextBox.Text.ToString().ToUpper();
                String courseName = CourseNameAddtextBox.Text.ToString();
                String courseCurriculum = CurriculumAddtextBox.Text.ToString();
                String courseDesc = DescriptionAddtextBox.Text.ToString();
                String coursePre = CoursePreAddtextBox.Text.ToString();
                String courseEqi = CourseEquivalentAddtextBox.Text.ToString();
                String courseCate = CourseCategoryAddComboBox.Text.ToString();
                int courseCredit = int.Parse(CourseCreditAddComboBox.Text.ToString());

                String query = "INSERT INTO Course VALUES('" + courseCode + "', '" + courseName + "', '" + courseDesc + "', '" + coursePre + "', '" + courseEqi + "', '" + courseCurriculum + "', '" + courseCate + "', " + courseCredit + ")";
                OleDbCommand cmd = new OleDbCommand(query, conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Data Saved!");
                CourseCodeAddtextBox.Text = "";
                CourseNameAddtextBox.Text = "";
                CurriculumAddtextBox.Text = "";
                DescriptionAddtextBox.Text = "";
                CoursePreAddtextBox.Text = "";
                CourseEquivalentAddtextBox.Text = "";
                CourseCategoryAddComboBox.Text = "";
                CourseCreditAddComboBox.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }
        }

        private void label3_Click_2(object sender, EventArgs e)
        {

        }

        private void CourseDesLabel_Click(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void CourseInfoSearchTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = CourseInfoSearchButton;
        }

        private void courseSearchorIDLabel_Click(object sender, EventArgs e)
        {

        }

        private void EditCourseInfoButton_Click(object sender, EventArgs e)
        {
            CourseNameTextBox.Enabled = true;
            CurriculumTextBox.Enabled = true;
            CourseDesTextBox8.Enabled = true;
            CourseEquivalentTextBox.Enabled = true;
            CourseCategoryComboBox.Enabled = true;
            CoursePreTextBox.Enabled = true;
            EditCourseInfoButton.Text = "Save";
            EditCourseInfoButton.Click += new EventHandler(SaveCourse_Click);
            this.AcceptButton = EditCourseInfoButton;
        }

        private void SaveCourse_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String courseCode = CourseCodeTextBox.Text.ToString();
                String courseName = CourseNameTextBox.Text.ToString();
                String courseCurriculum = CurriculumTextBox.Text.ToString();
                String courseDesc = CourseDesTextBox8.Text.ToString();
                String coursePre = CoursePreTextBox.Text.ToString();
                String courseEqi = CourseEquivalentTextBox.Text.ToString();
                String courseCate = CourseCategoryComboBox.Text.ToString();

                OleDbCommand cmd = new OleDbCommand("UPDATE Course SET [Course_Name] = '" + courseName + "', [Year_of_Syllabus] = '" + courseCurriculum + "' , [Course_Description] = '" + courseDesc + "', [Course_Pre-requisite] = '" + coursePre + "', [Course_Equivalent] = '" + courseEqi + "', [Course_Categories] = '" + courseCate + "' WHERE [Course_Code] = '" + courseCode + "'", conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Saved!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
                CourseCodeTextBox.Enabled = false;
                CourseNameTextBox.Enabled = false;
                CurriculumTextBox.Enabled = false;
                CourseDesTextBox8.Enabled = false;
                CourseEquivalentTextBox.Enabled = false;
                CourseCategoryComboBox.Enabled = false;
                CoursePreTextBox.Enabled = false;
                EditCourseInfoButton.Text = "Edit";
                EditCourseInfoButton.Click += new EventHandler(EditCourseInfoButton_Click);
            }
        }
        private void ClearCoursebutton_Click(object sender, EventArgs e)
        {
            if (CourseCodeAddtextBox.Text == "" && CourseNameAddtextBox.Text == "" &&
               CurriculumAddtextBox.Text == "" && DescriptionAddtextBox.Text == "" &&
               CoursePreTextBox.Text == "" && CourseEquivalentAddtextBox.Text == "" &&
               CourseCategoryAddComboBox.Text == "")
            {
                MessageBox.Show("Nothing to clear!");
            }
            else
            {
                CourseCodeAddtextBox.Text = "";
                CourseNameAddtextBox.Text = "";
                CurriculumAddtextBox.Text = "";
                DescriptionAddtextBox.Text = "";
                CoursePreAddtextBox.Text = "";
                CourseEquivalentAddtextBox.Text = "";
                CourseCategoryAddComboBox.Text = "";
            }
        }

        private void CourseInfoSearchButton_Click(object sender, EventArgs e)
        {
            int count;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String courseSearch = CourseInfoSearchTextBox.Text.ToString();

                String query = "SELECT * FROM Course WHERE Course_Code = '" + courseSearch + "'";
                String queryCount = "SELECT COUNT(*) FROM Course WHERE Course_Code = '" + courseSearch + "'";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbCommand cmd2 = new OleDbCommand(queryCount, conn);

                cmd.ExecuteNonQuery();

                OleDbDataReader reader = cmd.ExecuteReader();
                OleDbDataReader readCount = cmd2.ExecuteReader();

                while (readCount.Read())
                {
                    count = readCount.GetInt32(0);
                    if (count == 0)
                    {
                        MessageBox.Show("Result not found!");
                        CourseCodeTextBox.Text = "";
                        CourseNameTextBox.Text = "";
                        CurriculumTextBox.Text = "";
                        CourseDesTextBox8.Text = "";
                        CoursePreTextBox.Text = "";
                        CourseEquivalentTextBox.Text = "";
                        CourseCategoryComboBox.Text = "";
                        CourseCreditComboBox.Text = "";
                    }
                    else
                    {
                        while (reader.Read())
                        {
                            CourseCodeTextBox.Text = reader.GetString(0);
                            CourseNameTextBox.Text = reader.GetString(1);
                            CurriculumTextBox.Text = reader.GetString(5);
                            CourseDesTextBox8.Text = reader.GetString(2);
                            CoursePreTextBox.Text = reader.GetString(3);
                            CourseEquivalentTextBox.Text = reader.GetString(4);
                            CourseCategoryComboBox.Text = reader.GetString(6);
                            CourseCreditComboBox.Text = reader.GetInt32(7).ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }
        }

        private void EnterSearchTapButton_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(CousesTapPage);
        }

        private void EnterSearchTapButton_Click_1(object sender, EventArgs e)
        {
            tabControl1.SelectTab(CousesTapPage);
            CourseTabControl.SelectTab(CourseInfoTap);
        }

        private void CourseCodeAddtextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = AddCoursebutton;
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                int id = int.Parse(IdAddtextBox.Text.ToString());
                String firstname = FNameTextBox.Text.ToString();
                String lastname = SurAddTextBox.Text.ToString();
                String faculty = FacultyAddComboBox.Text.ToString();
                String major = MajorAddComboBox.Text.ToString();
                String status = StatusAddComboBox.Text.ToString();
                String degree = DegreeAddComboBox.Text.ToString();
                double GPA = double.Parse(GPAAddTextBox.Text.ToString());
                int year = int.Parse(YearAddTextBox.Text.ToString());
                String nationality = NationTextBox.Text.ToString();
                String religion = ReligionAddTextBox.Text.ToString();
                String address = AddressAddtextBox.Text.ToString();
                String phone = PhoneAddTextBox.Text.ToString();
                String email = EmailAddextBox.Text.ToString();
                String cur = CurriculumStudentComboBox.Text.ToString();
                int credit = int.Parse(CreditStudentTextBox.Text.ToString());

                String query = "INSERT INTO Student VALUES(" + id + ", '" + firstname + "', '" + lastname + "', '" + address + "', '" + faculty + "', '" + major + "', " + year + ", '" + status + "', " + GPA + ", '" + phone + "', '" + email + "', '" +nationality + "', '" + religion + "', '" + degree + "', '" + cur + "', " + credit + ")";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Data Saved!");
                IdAddtextBox.Text = "";
                FNameTextBox.Text = "";
                SurAddTextBox.Text = "";
                FacultyAddComboBox.Text = "";
                MajorAddComboBox.Text = "";
                StatusAddComboBox.Text = "";
                DegreeAddComboBox.Text = "";
                GPAAddTextBox.Text = "";
                YearAddTextBox.Text = "";
                NationTextBox.Text = "";
                ReligionAddTextBox.Text = "";
                AddressAddtextBox.Text = "";
                PhoneAddTextBox.Text = "";
                EmailAddextBox.Text = "";
                CurriculumStudentComboBox.Text = "";
                CreditStudentTextBox.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }



        }

        private void ViewButton_Click(object sender, EventArgs e)
        {
            if (FirstNameTextBox.Text == "")
            {
                MessageBox.Show("Search Student First!");
            }
            else
            {
                ((Control)this.CourseResultTap).Enabled = true;
                StudentTabControl.SelectTab(CourseResultTap);
                ShowStudentName.Text = FirstNameTextBox.Text + " " + SurnameTextBox.Text;
                ShowStudentID.Text = IDTextBox.Text;
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
                string strSql = "SELECT Course_Registration.Course_Code, Course.Course_Name FROM Course_Registration INNER JOIN Course ON Course_Registration.Course_Code = Course.Course_Code WHERE Student_ID = " + int.Parse(IDTextBox.Text.ToString());
                OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                conn.Open();
                cmdGrid.CommandType = CommandType.Text;
                OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                DataTable registered = new DataTable();
                da.Fill(registered);
                CourseRegistedDataGridView.DataSource = registered;
                this.CourseRegistedDataGridView.Sort(this.CourseRegistedDataGridView.Columns["Course_Code"], ListSortDirection.Ascending);

                string query2 = "SELECT Course.Course_Code, Course.Course_Name FROM Course INNER JOIN Student ON Course.Year_of_Syllabus = Student.Student_Syllabus OR Course.Year_of_Syllabus = 'All' WHERE Student_ID = " + int.Parse(IDTextBox.Text.ToString()) + " AND Course.Course_Code NOT IN (SELECT Course_Code FROM Course_Registration WHERE Student_ID = "+ int.Parse(IDTextBox.Text.ToString())+ ")";
                OleDbCommand cmdGrid2 = new OleDbCommand(query2, conn);
                cmdGrid2.CommandType = CommandType.Text;
                OleDbDataAdapter d2 = new OleDbDataAdapter(cmdGrid2);
                DataTable unregistered = new DataTable();
                d2.Fill(unregistered);
                dataGridView1.DataSource = unregistered;
                this.dataGridView1.Sort(this.dataGridView1.Columns["Course_Code"], ListSortDirection.Ascending);
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void LecturerAddbutton_Click(object sender, EventArgs e)
        {
                string lecturerName = comboBox3.Text;
                string courseName = label25.Text;
                string day = comboBox1.Text;
                string time = comboBox2.Text;
                int lecturerID, offerID;
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";

                try
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand("SELECT Lecturer_ID FROM Lecturer WHERE Lecturer_Name = '" + lecturerName + "'", conn);
                    OleDbDataReader readLecturerID = cmd.ExecuteReader();
                    while (readLecturerID.Read())
                    {
                        lecturerID = readLecturerID.GetInt32(0);
                        OleDbCommand cmd2 = new OleDbCommand("UPDATE Course_Offer SET [Lecturer_ID] = " + lecturerID + " WHERE [Course_Code] = '" + courseName + "'", conn);
                        cmd2.ExecuteNonQuery();

                        OleDbCommand cmdTimeTable = new OleDbCommand("SELECT Offer_ID FROM Course_Offer WHERE Course_Code = '" + courseName + "'", conn);
                        OleDbDataReader readOfferID = cmdTimeTable.ExecuteReader();
                        while (readOfferID.Read())
                        {
                            offerID = readOfferID.GetInt32(0);
                            OleDbCommand cmd3 = new OleDbCommand("INSERT INTO Timetable(Offer_ID, Course_Date, Course_Time, Lecturer_ID) VALUES (" + offerID + ", '" + day + "', '" + time + "', " + lecturerID + ")", conn);
                            cmd3.ExecuteNonQuery();
                        }
                    }
                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                    comboBox3.Enabled = false;
                    LecturerAddbutton.Text = "Edit";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                finally
                {
                    this.dataTable1TableAdapter1.Fill(this.dataSet2.DataTable1);
                    this.reportViewer2.RefreshReport();
                    MessageBox.Show("Saved!");
                    conn.Close();
                }
            
        }

        private void SelectToCouresesOffer_Click(object sender, EventArgs e)
        {

        }

        private void EditStudentButton_Click(object sender, EventArgs e)
        {
            FirstNameTextBox.Enabled = true;
            SurnameTextBox.Enabled = true;
            FacultyComboBox.Enabled = true;
            MajorComboBox.Enabled = true;
            StatusComboBox.Enabled = true;
            DegreeComboBox.Enabled = true;
            GPATextBox.Enabled = true;
            AcademicTextBox.Enabled = true;
            NationalityTextBox.Enabled = true;
            ReligionTextBox.Enabled = true;
            AddressTextBox.Enabled = true;
            PhonetextBox.Enabled = true;
            EmailTextBox.Enabled = true;
            CreditStudentShowTextBox.Enabled = true;
            CurriculumComboBox.Enabled = true;
            EditStudentButton.Text = "Save";
            EditStudentButton.Click += new EventHandler(SaveStudent_Click);
            this.AcceptButton = EditStudentButton;
        }
        private void SaveStudent_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                int id = int.Parse(IDTextBox.Text.ToString());
                String firstname = FirstNameTextBox.Text.ToString();
                String lastname = SurnameTextBox.Text.ToString();
                String faculty = FacultyComboBox.Text.ToString();
                String major = MajorComboBox.Text.ToString();
                String status = StatusComboBox.Text.ToString();
                String degree = DegreeComboBox.Text.ToString();
                double GPA = double.Parse(GPATextBox.Text.ToString());
                int year = int.Parse(AcademicTextBox.Text.ToString());
                String nationality = NationalityTextBox.Text.ToString();
                String religion = ReligionTextBox.Text.ToString();
                String curr = CurriculumComboBox.Text.ToString();
                int credit = int.Parse(CreditStudentShowTextBox.Text.ToString());
                String address = AddressTextBox.Text.ToString();
                String phone = PhonetextBox.Text.ToString();
                String email = EmailTextBox.Text.ToString();

                OleDbCommand cmd = new OleDbCommand("UPDATE Student SET [Student_First_Name] = '" + firstname + "' , [Student_Last_Name] = '" + lastname + "', [Student_Faculty] = '" + faculty + "', [Student_Major] = '" + major +
                    "', [Student_Status] = '" + status + "', [Student_Level] = '" + degree + "', [Student_GPA] = " + GPA + ", [Student_Year_of_Study] = " + year + ", [Student_Nationality] = '" + nationality + "', [Student_Religion] = '" + religion +
                    "', [Student_Address] = '" + address + "', [Student_Phone] = '" + phone + "', [Student_Email] = '" + email + "', [Student_Syllabus] = '" + curr + "', [Student_Credit] = " + credit + " WHERE [Student_ID] = " + id, conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Saved!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
                IDTextBox.Enabled = false;
                FirstNameTextBox.Enabled = false;
                SurnameTextBox.Enabled = false;
                FacultyComboBox.Enabled = false;
                MajorComboBox.Enabled = false;
                StatusComboBox.Enabled = false;
                DegreeComboBox.Enabled = false;
                GPATextBox.Enabled = false;
                AcademicTextBox.Enabled = false;
                NationalityTextBox.Enabled = false;
                ReligionTextBox.Enabled = false;
                AddressTextBox.Enabled = false;
                PhonetextBox.Enabled = false;
                EmailTextBox.Enabled = false;
                CurriculumComboBox.Enabled = false;
                CreditStudentShowTextBox.Enabled = false;
                EditStudentButton.Text = "Edit";
                EditStudentButton.Click += new EventHandler(EditStudentButton_Click);
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {

            if (IdAddtextBox.Text == "" && FNameTextBox.Text == "" &&
               SurAddTextBox.Text == "" && FacultyAddComboBox.Text == "" &&
              MajorAddComboBox.Text == "" && StatusAddComboBox.Text == "" &&
               DegreeAddComboBox.Text == "" && GPAAddTextBox.Text == "" &&
                YearAddTextBox.Text == "" && NationTextBox.Text == "" &&
                ReligionAddTextBox.Text == "" && AddressAddtextBox.Text == "" &&
                PhoneAddTextBox.Text == "" && EmailAddextBox.Text == "")
            {
                MessageBox.Show("Nothing to clear!");
            }
            else
            {
                IdAddtextBox.Text = "";
                FNameTextBox.Text = "";
                SurAddTextBox.Text = "";
                FacultyAddComboBox.Text = "";
                MajorAddComboBox.Text = "";
                StatusAddComboBox.Text = "";
                DegreeAddComboBox.Text = "";
                GPAAddTextBox.Text = "";
                YearAddTextBox.Text = "";
                NationTextBox.Text = "";
                ReligionAddTextBox.Text = "";
                AddressAddtextBox.Text = "";
                PhoneAddTextBox.Text = "";
                EmailAddextBox.Text = "";

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int count;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String studentSearch = StudentSearchTextbox.Text.ToString();

                String query = "SELECT * FROM Student WHERE Student_First_Name = '" + studentSearch + "'";
                String queryCount = "SELECT COUNT(*) FROM Student WHERE Student_First_Name = '" + studentSearch + "'";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbCommand cmd2 = new OleDbCommand(queryCount, conn);

                cmd.ExecuteNonQuery();

                OleDbDataReader reader = cmd.ExecuteReader();
                OleDbDataReader readCount = cmd2.ExecuteReader();

                
                while (readCount.Read())
                {
                    count = readCount.GetInt32(0);
                    if (count == 0)
                    {
                        MessageBox.Show("Result not found!");
                        IDTextBox.Text = "";
                        FirstNameTextBox.Text = "";
                        SurnameTextBox.Text = "";
                        FacultyComboBox.Text = "";
                        MajorComboBox.Text = "";
                        StatusComboBox.Text = "";
                        DegreeComboBox.Text = "";
                        GPATextBox.Text = "";
                        CurriculumComboBox.Text = "";
                        NationalityTextBox.Text = "";
                        ReligionTextBox.Text = "";
                        AddressTextBox.Text = "";
                        PhonetextBox.Text = "";
                        EmailTextBox.Text = "";
                        CreditStudentShowTextBox.Text = "";
                        AcademicTextBox.Text = "";
                    }
                    else
                    {
                        while (reader.Read())
                        {

                            IDTextBox.Text = reader.GetInt32(0).ToString();
                            FirstNameTextBox.Text = reader.GetString(1);
                            SurnameTextBox.Text = reader.GetString(2);
                            FacultyComboBox.Text = reader.GetString(4);
                            MajorComboBox.Text = reader.GetString(5);
                            StatusComboBox.Text = reader.GetString(7);
                            DegreeComboBox.Text = reader.GetString(13);
                            AcademicTextBox.Text = reader.GetInt32(6).ToString();
                            NationTextBox.Text = reader.GetString(11);
                            ReligionTextBox.Text = reader.GetString(12);
                            AddressTextBox.Text = reader.GetString(3);
                            PhonetextBox.Text = reader.GetString(9);
                            EmailTextBox.Text = reader.GetString(10);
                            GPATextBox.Text = reader.GetDouble(8).ToString();
                            CurriculumComboBox.Text = reader.GetString(14);
                            CreditStudentShowTextBox.Text = reader.GetInt32(15).ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }


        }

        private void FacultyAddComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (FacultyAddComboBox.SelectedIndex == 0)
            {
                MajorAddComboBox.Items.Clear();
                MajorAddComboBox.Items.Add("IHM");
                MajorAddComboBox.Items.Add("IBM");
                MajorAddComboBox.Items.Add("ABM");
                MajorAddComboBox.Items.Add("MKT");
            }
            else if (FacultyAddComboBox.SelectedIndex == 1)
            {
                MajorAddComboBox.Items.Clear();
                MajorAddComboBox.Items.Add("ISM");
                MajorAddComboBox.Items.Add("Computer Animation");
                MajorAddComboBox.Items.Add("IT");
            }
            else if (FacultyAddComboBox.SelectedIndex == 2)
            {
                MajorAddComboBox.Items.Clear();
                MajorAddComboBox.Items.Add("Broadcast");
                MajorAddComboBox.Items.Add("Com Arts");
                MajorAddComboBox.Items.Add("Advertising");
            }
            else
            {
                MajorAddComboBox.Items.Clear();
                MajorAddComboBox.Items.Add("CMD");
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (FacultyComboBox.SelectedIndex == 0)
            {
                MajorComboBox.Items.Clear();
                MajorComboBox.Items.Add("IHM");
                MajorComboBox.Items.Add("IBM");
                MajorComboBox.Items.Add("ABM");
                MajorComboBox.Items.Add("MKT");
            }
            else if (FacultyComboBox.SelectedIndex == 1)
            {
                MajorComboBox.Items.Clear();
                MajorComboBox.Items.Add("ISM");
                MajorComboBox.Items.Add("Computer Animation");
                MajorComboBox.Items.Add("IT");
            }
            else if (FacultyComboBox.SelectedIndex == 2)
            {
                MajorComboBox.Items.Clear();
                MajorComboBox.Items.Add("Broadcast");
                MajorComboBox.Items.Add("Com Arts");
                MajorComboBox.Items.Add("Advertising");
            }
            else
            {
                MajorComboBox.Items.Clear();
                MajorComboBox.Items.Add("CMD");
            }
        }

        private void CourseRegistedDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void AddLecturerButton_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String lecturerName = AddLecturerNameTextBox.Text.ToString();
                String lecturerType = AddLecturerTypeComboBox.Text.ToString();

                String query = "INSERT INTO Lecturer (Lecturer_Name, Lecturer_Type) VALUES('" + lecturerName + "', '" + lecturerType + "')";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Data Saved!");
                AddLecturerNameTextBox.Text = "";
                AddLecturerTypeComboBox.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }
        }

        private void groupBox3_Enter_1(object sender, EventArgs e)
        {

        }

        private void AddCurriculumAddButton_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                int year = int.Parse(AddCurriculumYearTextBox.Text.ToString());

                String query = "INSERT INTO Curriculum (Curriculum_Year) VALUES(" + year + ")";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Data Saved!");
                AddCurriculumYearTextBox.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }
        }

        private void IdAddtextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = AddButton;
            string actualdata = string.Empty;
            char[] entereddata = IdAddtextBox.Text.ToCharArray();
            foreach (char aChar in entereddata.AsEnumerable())
            {
                if (Char.IsDigit(aChar))
                {
                    actualdata = actualdata + aChar;
                }
                else
                {
                    MessageBox.Show("Please Enter in Numeric Format");
                    actualdata.Replace(aChar, ' ');
                    actualdata.Trim();
                }
            }
            IdAddtextBox.Text = actualdata;
            IdAddtextBox.MaxLength = 9;
        }

        private void MajorComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void StudentSearchTextbox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = StudentSearchButton;
        }

        private void CourseEnrollConfirmButton_Click(object sender, EventArgs e)
        {
            String CourseCode = CourseEnrollTextBox.Text.ToString().ToUpper();
            int StudentID = int.Parse(ShowStudentID.Text.ToString());

            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String queryDuplicate = "SELECT COUNT(*) FROM Course_Registration WHERE Student_ID = " + StudentID + " AND Course_Code = '" + CourseCode + "'";
                OleDbCommand cmdDuplicate = new OleDbCommand(queryDuplicate, conn);
                OleDbDataReader reader = cmdDuplicate.ExecuteReader();
                int count;
                while (reader.Read())
                {
                    count = reader.GetInt32(0);
                    if (count == 0)
                    {
                        try
                        {
                            String query = "INSERT INTO Course_Registration (Student_ID, Course_Code, Course_Status) VALUES( " + StudentID + ", '" + CourseCode + "', 'Taken' )";

                            OleDbCommand cmd = new OleDbCommand(query, conn);
                            cmd.ExecuteNonQuery();

                            MessageBox.Show("Data Saved!");
                            CourseEnrollTextBox.Text = "";
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Duplicate Enrollment!");
                        CourseEnrollTextBox.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                try
                {
                    string strSql = "SELECT Course_Registration.Course_Code, Course.Course_Name FROM Course_Registration INNER JOIN Course ON Course_Registration.Course_Code = Course.Course_Code WHERE Student_ID = " + int.Parse(IDTextBox.Text.ToString());
                    OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                    cmdGrid.CommandType = CommandType.Text;
                    OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                    DataTable registered = new DataTable();
                    da.Fill(registered);
                    CourseRegistedDataGridView.DataSource = registered;
                    this.CourseRegistedDataGridView.Sort(this.CourseRegistedDataGridView.Columns["Course_Code"], ListSortDirection.Ascending);

                    string query2 = "SELECT Course.Course_Code, Course.Course_Name FROM Course INNER JOIN Student ON Course.Year_of_Syllabus = Student.Student_Syllabus OR Course.Year_of_Syllabus = 'All' WHERE Student_ID = " + int.Parse(IDTextBox.Text.ToString()) + " AND Course.Course_Code NOT IN (SELECT Course_Code FROM Course_Registration WHERE Student_ID = " + int.Parse(IDTextBox.Text.ToString()) + ")";
                    OleDbCommand cmdGrid2 = new OleDbCommand(query2, conn);
                    cmdGrid2.CommandType = CommandType.Text;
                    OleDbDataAdapter d2 = new OleDbDataAdapter(cmdGrid2);
                    DataTable unregistered = new DataTable();
                    d2.Fill(unregistered);
                    dataGridView1.DataSource = unregistered;
                    this.dataGridView1.Sort(this.dataGridView1.Columns["Course_Code"], ListSortDirection.Ascending);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                conn.Close();
            }
        }

        private void IDTextBox_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void CourseEnrollTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = CourseEnrollConfirmButton;
        }

        private void CourseEnrollCancelButton_Click(object sender, EventArgs e)
        {
            CourseEnrollTextBox.Text = "";
        }

        private void EditLecturerButton_Click(object sender, EventArgs e)
        {
            ModifyLecturerNameTextBox.Enabled = true;
            ModifyLecturerTypeComboBox.Enabled = true;
            EditLecturerButton.Text = "Save";
            EditLecturerButton.Click += new EventHandler(SaveLecturer_Click);
            this.AcceptButton = EditLecturerButton;
        }
        private void SaveLecturer_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String name = ModifyLecturerNameTextBox.Text.ToString();
                String type = ModifyLecturerTypeComboBox.Text.ToString();

                OleDbCommand cmd = new OleDbCommand("UPDATE Lecturer SET [Lecturer_Name] = '" + name + "' , [Lecturer_Type] = '" + type + "'", conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Saved!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
                ModifyLecturerNameTextBox.Enabled = false;
                ModifyLecturerTypeComboBox.Enabled = false;
                EditLecturerButton.Text = "Edit";
                EditLecturerButton.Click += new EventHandler(EditLecturerButton_Click);
            }
        }
        private void RemoveLecturerSearchButton_Click(object sender, EventArgs e)
        {
            int count;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String lecturerSearch = RemoveLecturerSearchTextBox.Text.ToString();

                String query = "SELECT * FROM Lecturer WHERE Lecturer_Name = '" + lecturerSearch + "'";
                String queryCount = "SELECT COUNT(*) FROM Lecturer WHERE Lecturer_Name = '" + lecturerSearch + "'";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbCommand cmd2 = new OleDbCommand(queryCount, conn);

                cmd.ExecuteNonQuery();

                OleDbDataReader reader = cmd.ExecuteReader();
                OleDbDataReader readCount = cmd2.ExecuteReader();


                while (readCount.Read())
                {
                    count = readCount.GetInt32(0);
                    if (count == 0)
                    {
                        MessageBox.Show("Result not found!");
                        ShowLecturerNameRemove.Text = "";
                        ShowLecturerTypeRemove.Text = "";
                    }
                    else
                    {
                        while (reader.Read())
                        {
                            ShowLecturerNameRemove.Text = reader.GetString(1);
                            ShowLecturerTypeRemove.Text = reader.GetString(2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }
        }

        private void ClearAddLecturerButton_Click(object sender, EventArgs e)
        {
            AddLecturerNameTextBox.Text = "";
            AddLecturerTypeComboBox.Text = "";
        }

        private void SearchModifyLecturerButton_Click(object sender, EventArgs e)
        {
            int count;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String lecturerSearch = SearchModifyLecturerTextBox.Text.ToString();

                String query = "SELECT * FROM Lecturer WHERE Lecturer_Name = '" + lecturerSearch + "'";
                String queryCount = "SELECT COUNT(*) FROM Lecturer WHERE Lecturer_Name = '" + lecturerSearch + "'";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbCommand cmd2 = new OleDbCommand(queryCount, conn);

                cmd.ExecuteNonQuery();

                OleDbDataReader reader = cmd.ExecuteReader();
                OleDbDataReader readCount = cmd2.ExecuteReader();


                while (readCount.Read())
                {
                    count = readCount.GetInt32(0);
                    if (count == 0)
                    {
                        MessageBox.Show("Result not found!");
                        ModifyLecturerNameTextBox.Text = "";
                        ModifyLecturerTypeComboBox.Text = "";
                    }
                    else
                    {
                        while (reader.Read())
                        {
                            ModifyLecturerNameTextBox.Text = reader.GetString(1);
                            ModifyLecturerTypeComboBox.Text = reader.GetString(2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }   
        }

        private void SearchModifyLecturerTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = SearchModifyLecturerButton;
        }

        private void RemoveLecturerSearchTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = RemoveLecturerSearchButton;
        }

        private void CancelRemoveLecturerButton_Click(object sender, EventArgs e)
        {
            ShowLecturerNameRemove.Text = "";
            ShowLecturerTypeRemove.Text = "";
        }

        private void RemoveLecturerButton_Click(object sender, EventArgs e)
        {
            if((MessageBox.Show("Are you sure ?", "Delete Lecturer", 
    MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes))
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
                String name = ShowLecturerNameRemove.Text.ToString();
                try
                {
                    conn.Open();
                    String query = "DELETE FROM Lecturer WHERE Lecturer_Name = '" + name + "'";
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                finally
                {
                    conn.Close();
                    RemoveLecturerSearchTextBox.Text = "";
                    ShowLecturerNameRemove.Text = "";
                    ShowLecturerTypeRemove.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Nothing Delete");
            }
        }

        private void SearchStudentListButton_Click(object sender, EventArgs e)
        {
            String StudentName = SearchStudentTextBox.Text.ToString();

            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                string strSql;
                if (StudentName == "")
                {
                    strSql = "SELECT * FROM Student";
                }
                else
                {
                    strSql = "SELECT * FROM Student WHERE Student_First_Name= '" + StudentName + "'";
                }
                OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                cmdGrid.CommandType = CommandType.Text;
                OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                DataTable registered = new DataTable();
                da.Fill(registered);
                StudentsListDataGridView.DataSource = registered;
                this.StudentsListDataGridView.Sort(this.StudentsListDataGridView.Columns["Student_ID"], ListSortDirection.Ascending);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }
        }

        private void SearchStudentTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = SearchStudentListButton;
            SearchStudentListButton.Text = "Search";
        }

        private void CourseResultTap_Click(object sender, EventArgs e)
        {

        }

        private void CourseListSearhTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = CourseListSearhbutton;
            if (CourseListSearhTextBox.Text == "")
            {

            }
            CourseListSearhbutton.Text = "Search";
        }

        private void CourseListSearhbutton_Click(object sender, EventArgs e)
        {
            String CourseCode = CourseListSearhTextBox.Text.ToString();

            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                string strSql;
                if (CourseCode == "")
                {
                    strSql = "SELECT Course_Code AS [Course Code], Course_Name AS [Course Name], Course_Description AS [Description], [Course_Pre-requisite] AS [Pre-requisite], Course_Equivalent AS [Equivalent], Year_of_Syllabus AS [Curriculum], Course_Categories AS [Categories], Credit FROM Course";
                }
                else
                {
                    strSql = "SELECT Course_Code AS [Course Code], Course_Name AS [Course Name], Course_Description AS [Description], [Course_Pre-requisite] AS [Pre-requisite], Course_Equivalent AS [Equivalent], Year_of_Syllabus AS [Curriculum], Course_Categories AS [Categories], Credit FROM Course WHERE Course_Code = '" + CourseCode + "'";
                }
                OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                cmdGrid.CommandType = CommandType.Text;
                OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                DataTable registered = new DataTable();
                da.Fill(registered);
                CourseListdataGridView.DataSource = registered;
                this.CourseListdataGridView.Sort(this.CourseListdataGridView.Columns["Course Code"], ListSortDirection.Ascending);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }
        }

        private void AddCurriculumClearButton_Click(object sender, EventArgs e)
        {
            AddCurriculumYearTextBox.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int count;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String curriculumSearch = ModifyCurriculumSearchTextBox.Text.ToString();

                String query = "SELECT * FROM Curriculum WHERE Curriculum_Year = '" + curriculumSearch + "'";
                String queryCount = "SELECT COUNT(*) FROM Curriculum WHERE Curriculum_Year = '" + curriculumSearch + "'";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbCommand cmd2 = new OleDbCommand(queryCount, conn);

                cmd.ExecuteNonQuery();

                OleDbDataReader reader = cmd.ExecuteReader();
                OleDbDataReader readCount = cmd2.ExecuteReader();


                while (readCount.Read())
                {
                    count = readCount.GetInt32(0);
                    if (count == 0)
                    {
                        MessageBox.Show("Result not found!");
                        ModifyCurriculumTextBox.Text = "";
                    }
                    else
                    {
                        while (reader.Read())
                        {
                            ModifyCurriculumTextBox.Text = reader.GetString(1);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }   
        }

        private void ModifyCurriculumSearchTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = ModifyCurriculumSearchButton;
        }

        private void ModifyCurriculumButton_Click(object sender, EventArgs e)
        {
            ModifyCurriculumTextBox.Enabled = true;
            ModifyCurriculumButton.Text = "Save";
            ModifyCurriculumButton.Click += new EventHandler(SaveCurriculum_Click);
            this.AcceptButton = EditLecturerButton;
        }
        private void SaveCurriculum_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+appPath+"\\CourseManagementDatabase.accdb";
            String year = ModifyCurriculumSearchTextBox.Text.ToString();
            int count, ID;
            try
            {
                conn.Open();
                MessageBox.Show(year);
                String query = "SELECT * FROM Curriculum WHERE Curriculum_Year = '" + year + "'";
                String queryCount = "SELECT COUNT(*) FROM Curriculum WHERE Curriculum_Year = '" + year + "'";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbCommand cmd2 = new OleDbCommand(queryCount, conn);
            
                cmd.ExecuteNonQuery();

                OleDbDataReader reader = cmd.ExecuteReader();
                OleDbDataReader readCount = cmd2.ExecuteReader();
                    
                while (readCount.Read())
                {
                    count = readCount.GetInt32(0);
                    if (count == 0)
                    {
                        MessageBox.Show("Result not found!");
                    }
                    else
                    {
                        while (reader.Read())
                        {
                            ID = reader.GetInt32(0);
                            String newYear = ModifyCurriculumTextBox.Text.ToString();
                            MessageBox.Show(ID.ToString());
                            OleDbCommand cmd3 = new OleDbCommand("UPDATE Curriculum SET [Curriculum_Year] = '" + newYear + "' WHERE [Curriculum_ID] = " + ID, conn);
                            cmd3.ExecuteNonQuery();
                            MessageBox.Show("Saved!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
                ModifyCurriculumTextBox.Enabled = false;
                ModifyCurriculumButton.Text = "Edit";
                ModifyCurriculumButton.Click += new EventHandler(ModifyCurriculumButton_Click);
            }
        }

        private void ModifyCurriculumTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = ModifyCurriculumButton;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            ModifyCurriculumTextBox.Text = "";
            ModifyCurriculumSearchTextBox.Text = "";
        }

        private void ModifyCurriculumGroupBox_Enter(object sender, EventArgs e)
        {

        }

        private void RemoveCurriculumSearchButton_Click(object sender, EventArgs e)
        {
            int count;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                String curriculumSearch = RemoveCurriculumSearchTextBox.Text.ToString();

                String query = "SELECT * FROM Curriculum WHERE Curriculum_Year = '" + curriculumSearch + "'";
                String queryCount = "SELECT COUNT(*) FROM Curriculum WHERE Curriculum_Year = '" + curriculumSearch + "'";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbCommand cmd2 = new OleDbCommand(queryCount, conn);

                cmd.ExecuteNonQuery();

                OleDbDataReader reader = cmd.ExecuteReader();
                OleDbDataReader readCount = cmd2.ExecuteReader();


                while (readCount.Read())
                {
                    count = readCount.GetInt32(0);
                    if (count == 0)
                    {
                        MessageBox.Show("Result not found!");
                    }
                    else
                    {
                        while (reader.Read())
                        {
                            RemoveCurriculumLabel.Text = reader.GetString(1);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                conn.Close();
            }   
        }

        private void button3_Click(object sender, EventArgs e)
        {
            RemoveCurriculumSearchTextBox.Text = "";
            RemoveCurriculumLabel.Text = "";
        }

        private void RemoveCurriculumButton_Click(object sender, EventArgs e)
        {
            if ((MessageBox.Show("Are you sure ?", "Delete Curriculum",
    MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes))
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";
                String curriculum = RemoveCurriculumLabel.Text.ToString();
                try
                {
                    conn.Open();
                    String query = "DELETE FROM Curriculum WHERE Curriculum_Year = '" + curriculum + "'";
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                finally
                {
                    conn.Close();
                    RemoveCurriculumLabel.Text = "";
                    RemoveCurriculumSearchTextBox.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Nothing Delete");
            }
        }

        private void RemoveCurriculumSearchTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = RemoveCurriculumSearchButton;
        }

        private void CourseManageTap_Click(object sender, EventArgs e)
        {

        }

        private void CourseSelectTextBox_TextChanged(object sender, EventArgs e)
        {
            this.AcceptButton = AddSelectCoursesButton;
        }

        private void RemoveStudentButton_Click(object sender, EventArgs e)
        {
            if ((MessageBox.Show("Are you sure ?", "Delete Curriculum",
    MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes))
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";
                String ID = IDTextBox.Text.ToString();
                try
                {
                    conn.Open();
                    String query = "DELETE FROM Student WHERE Student_ID = " + ID;
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Deleted Successful");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                finally
                {
                    conn.Close();
                    FirstNameTextBox.Text = "";
                    SurnameTextBox.Text = "";
                    FacultyComboBox.Text = "";
                    MajorComboBox.Text = "";
                    StatusComboBox.Text = "";
                    DegreeComboBox.Text = "";
                    GPATextBox.Text = "";
                    AcademicTextBox.Text = "";
                    NationalityTextBox.Text = "";
                    ReligionTextBox.Text = "";
                    AddressTextBox.Text = "";
                    PhonetextBox.Text = "";
                    EmailTextBox.Text = "";
                    CreditStudentShowTextBox.Text = "";
                    CurriculumComboBox.Text = "";
                    IDTextBox.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Nothing Delete");
            }
        }

        private void RemoveCourseButton_Click(object sender, EventArgs e)
        {
            if ((MessageBox.Show("Are you sure ?", "Delete Curriculum",
    MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes))
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";
                String courseCode = CourseCodeTextBox.Text.ToString();
                try
                {
                    conn.Open();
                    String query = "DELETE FROM Course WHERE Course_Code = '" + courseCode + "'";
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Deleted Successful");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                finally
                {
                    conn.Close();
                    CourseNameTextBox.Text = "";
                    CurriculumTextBox.Text = "";
                    CourseDesTextBox8.Text = "";
                    CourseEquivalentTextBox.Text = "";
                    CourseCategoryComboBox.Text = "";
                    CoursePreTextBox.Text = "";
                    CourseCodeTextBox.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Nothing Delete");
            }
        }

        private void StudentSearchReporttextBox_TextChanged(object sender, EventArgs e)
        {
        }

        private void StudentViewReportButton_Click(object sender, EventArgs e)
        {
            this.dataTable1TableAdapter.Fill(this.dataSet1.DataTable1);
            this.reportViewer1.RefreshReport();
        }

        private void AddSelectCoursesButton_Click(object sender, EventArgs e)
        {
            string offer = CoursesSelectdataGridView3.CurrentRow.Cells[0].FormattedValue.ToString();
            String strSql;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                strSql = "INSERT INTO Course_Offer (Course_Code) VALUES ('" + offer + "')";
                OleDbCommand cmd = new OleDbCommand(strSql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            try
            {
                strSql = "SELECT Course_Code FROM Course_Offer";
                OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                cmdGrid.CommandType = CommandType.Text;
                OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                DataTable registered = new DataTable();
                da.Fill(registered);
                dataGridView2.DataSource = registered;
                this.dataGridView2.Sort(this.dataGridView2.Columns["Course_Code"], ListSortDirection.Ascending);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                try
                {
                    strSql = "SELECT Course_Offer.Course_Code, Course.Course_Name FROM Course_Offer INNER JOIN Course ON Course_Offer.Course_Code = Course.Course_Code";
                    OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                    cmdGrid.CommandType = CommandType.Text;
                    OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                    DataTable registered = new DataTable();
                    da.Fill(registered);
                    CoursesOfferListdataGridView.DataSource = registered;
                    this.CoursesOfferListdataGridView.Sort(this.CoursesOfferListdataGridView.Columns["Course_Code"], ListSortDirection.Ascending);
                    CoursesOfferListdataGridView.CellMouseClick += new DataGridViewCellMouseEventHandler(CourseNameClick);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
            }
            conn.Close();
        }
        private void CourseNameClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";

                ((Control)this.Course_Detail_Tab).Enabled = true;
                CoursesOfferTabControl.SelectTab(Course_Detail_Tab);
                courseManageButton.Enabled = true;

                string course_code = CoursesOfferListdataGridView.CurrentRow.Cells[0].FormattedValue.ToString();
                string course_name = CoursesOfferListdataGridView.CurrentRow.Cells[1].FormattedValue.ToString();

                coursedetail_CourseCode.Text = course_code;
                coursedetail_CourseName.Text = course_name;
                try {
                    conn.Open();
                    string query2 = "SELECT DISTINCT(Course_Registration.Student_ID), Student.Student_First_Name FROM Course_Registration INNER JOIN Student ON Student.Student_ID = Course_Registration.Student_ID WHERE Course_Registration.Course_Code = '" + course_code + "'";

                    OleDbCommand cmdGrid2 = new OleDbCommand(query2, conn);
                    cmdGrid2.CommandType = CommandType.Text;
                    OleDbDataAdapter d2 = new OleDbDataAdapter(cmdGrid2);
                    DataTable registered = new DataTable();
                    d2.Fill(registered);
                    StudentCompleteCourseDetailData.DataSource = registered;
                    this.StudentCompleteCourseDetailData.Sort(this.StudentCompleteCourseDetailData.Columns["Student_ID"], ListSortDirection.Descending);

                    string query = "SELECT COUNT (*) FROM (SELECT DISTINCT(Course_Registration.Student_ID) FROM Course_Registration INNER JOIN Student ON Student.Student_ID = Course_Registration.Student_ID WHERE Course_Registration.Course_Code = '" + course_code + "')";

                    OleDbCommand cmdCount = new OleDbCommand(query, conn);
                    OleDbDataReader reader = cmdCount.ExecuteReader();
                    while (reader.Read())
                    {
                        label27.Text = reader.GetInt32(0).ToString();
                    }

                    string query4 = "SELECT DISTINCT (Student_ID), Student.Student_First_Name FROM Student INNER JOIN Course ON Student.Student_Syllabus = Course.Year_of_Syllabus OR Course.Year_of_Syllabus = 'All' WHERE Student_ID not in (SELECT Student_ID FROM Course_Registration WHERE Course_Code='"+course_code+"') AND Course_Code = '" + course_code + "'";

                    OleDbCommand cmdGrid3 = new OleDbCommand(query4, conn);
                    cmdGrid2.CommandType = CommandType.Text;
                    OleDbDataAdapter d3 = new OleDbDataAdapter(cmdGrid3);
                    DataTable unregistered = new DataTable();
                    d3.Fill(unregistered);
                    StudentInCompleteCourseDetailData.DataSource = unregistered;
                    this.StudentInCompleteCourseDetailData.Sort(this.StudentInCompleteCourseDetailData.Columns["Student_ID"], ListSortDirection.Descending);

                    string query5 = "SELECT COUNT (*) FROM (SELECT DISTINCT (Student_ID), Student.Student_First_Name FROM Student INNER JOIN Course ON Student.Student_Syllabus = Course.Year_of_Syllabus OR Course.Year_of_Syllabus = 'All' WHERE Student_ID not in (SELECT Student_ID FROM Course_Registration WHERE Course_Code='" + course_code + "') AND Course_Code = '" + course_code + "') ";

                    OleDbCommand cmdCount2 = new OleDbCommand(query5, conn);
                    OleDbDataReader reader2 = cmdCount2.ExecuteReader();
                    while (reader2.Read())
                    {
                        label28.Text = reader2.GetInt32(0).ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                finally
                {
                    conn.Close();
                }
            }
            else
            {
                
            }         
        }
        private void CourseNameClick1(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";

                ((Control)this.Course_Detail_Tab).Enabled = true;
                CoursesOfferTabControl.SelectTab(Course_Detail_Tab);
                courseManageButton.Enabled = false;
                
                string course_code = CoursesSelectdataGridView3.CurrentRow.Cells[0].FormattedValue.ToString();
                string course_name = CoursesSelectdataGridView3.CurrentRow.Cells[1].FormattedValue.ToString();

                label25.Text = (course_code + " " + course_name);
                coursedetail_CourseCode.Text = course_code;
                coursedetail_CourseName.Text = course_name;
                try
                {
                    conn.Open();
                    string query2 = "SELECT DISTINCT(Course_Registration.Student_ID), Student.Student_First_Name FROM Course_Registration INNER JOIN Student ON Student.Student_ID = Course_Registration.Student_ID WHERE Course_Registration.Course_Code = '" + course_code + "'";

                    OleDbCommand cmdGrid2 = new OleDbCommand(query2, conn);
                    cmdGrid2.CommandType = CommandType.Text;
                    OleDbDataAdapter d2 = new OleDbDataAdapter(cmdGrid2);
                    DataTable registered = new DataTable();
                    d2.Fill(registered);
                    StudentCompleteCourseDetailData.DataSource = registered;
                    this.StudentCompleteCourseDetailData.Sort(this.StudentCompleteCourseDetailData.Columns["Student_ID"], ListSortDirection.Descending);

                    string query = "SELECT COUNT (*) FROM (SELECT DISTINCT(Course_Registration.Student_ID) FROM Course_Registration INNER JOIN Student ON Student.Student_ID = Course_Registration.Student_ID WHERE Course_Registration.Course_Code = '" + course_code + "')";

                    OleDbCommand cmdCount = new OleDbCommand(query, conn);
                    OleDbDataReader reader = cmdCount.ExecuteReader();
                    while (reader.Read())
                    {
                        label27.Text = reader.GetInt32(0).ToString();
                    }

                    string query4 = "SELECT DISTINCT (Student_ID), Student.Student_First_Name FROM Student INNER JOIN Course ON Student.Student_Syllabus = Course.Year_of_Syllabus OR Course.Year_of_Syllabus = 'All' WHERE Student_ID not in (SELECT Student_ID FROM Course_Registration WHERE Course_Code='" + course_code + "') AND Course_Code = '" + course_code + "'";

                    OleDbCommand cmdGrid3 = new OleDbCommand(query4, conn);
                    cmdGrid2.CommandType = CommandType.Text;
                    OleDbDataAdapter d3 = new OleDbDataAdapter(cmdGrid3);
                    DataTable unregistered = new DataTable();
                    d3.Fill(unregistered);
                    StudentInCompleteCourseDetailData.DataSource = unregistered;
                    this.StudentInCompleteCourseDetailData.Sort(this.StudentInCompleteCourseDetailData.Columns["Student_ID"], ListSortDirection.Descending);

                    string query5 = "SELECT COUNT (*) FROM (SELECT DISTINCT (Student_ID), Student.Student_First_Name FROM Student INNER JOIN Course ON Student.Student_Syllabus = Course.Year_of_Syllabus OR Course.Year_of_Syllabus = 'All' WHERE Student_ID not in (SELECT Student_ID FROM Course_Registration WHERE Course_Code='" + course_code + "') AND Course_Code = '" + course_code + "') ";

                    OleDbCommand cmdCount2 = new OleDbCommand(query5, conn);
                    OleDbDataReader reader2 = cmdCount2.ExecuteReader();
                    while (reader2.Read())
                    {
                        label28.Text = reader2.GetInt32(0).ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
                finally
                {
                    conn.Close();
                }
            }
            else
            {

            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            String strSql;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                strSql = "DELETE * FROM Course_Offer";
                OleDbCommand cmd = new OleDbCommand(strSql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            try
            {
                strSql = "DELETE * FROM TimeTable";
                OleDbCommand cmd = new OleDbCommand(strSql, conn);
                cmd.ExecuteNonQuery();

                strSql = "SELECT Course_Code FROM Course_Offer";
                OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                cmdGrid.CommandType = CommandType.Text;
                OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                DataTable registered = new DataTable();
                da.Fill(registered);
                dataGridView2.DataSource = registered;
                this.dataGridView2.Sort(this.dataGridView2.Columns["Course_Code"], ListSortDirection.Ascending);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                try
                {
                    strSql = "SELECT Course_Offer.Course_Code, Course.Course_Name FROM Course_Offer INNER JOIN Course ON Course_Offer.Course_Code = Course.Course_Code";
                    OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                    cmdGrid.CommandType = CommandType.Text;
                    OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                    DataTable registered = new DataTable();
                    da.Fill(registered);
                    CoursesOfferListdataGridView.DataSource = registered;
                    this.CoursesOfferListdataGridView.Sort(this.CoursesOfferListdataGridView.Columns["Course_Code"], ListSortDirection.Ascending);                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
            }
            conn.Close();  
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            string offer = dataGridView2.CurrentRow.Cells[0].FormattedValue.ToString();
            String strSql;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";
            
            try
            {
                conn.Open();
                strSql = "DELETE FROM Course_Offer WHERE Course_Code = '" + offer + "'";
                OleDbCommand cmd = new OleDbCommand(strSql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            try
            {
                strSql = "SELECT Course_Code FROM Course_Offer";
                OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                cmdGrid.CommandType = CommandType.Text;
                OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                DataTable registered = new DataTable();
                da.Fill(registered);
                dataGridView2.DataSource = registered;
                this.dataGridView2.Sort(this.dataGridView2.Columns["Course_Code"], ListSortDirection.Ascending);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
            finally
            {
                try
                {
                    strSql = "SELECT Course_Offer.Course_Code, Course.Course_Name FROM Course_Offer INNER JOIN Course ON Course_Offer.Course_Code = Course.Course_Code";
                    OleDbCommand cmdGrid = new OleDbCommand(strSql, conn);
                    cmdGrid.CommandType = CommandType.Text;
                    OleDbDataAdapter da = new OleDbDataAdapter(cmdGrid);
                    DataTable registered = new DataTable();
                    da.Fill(registered);
                    CoursesOfferListdataGridView.DataSource = registered;
                    this.CoursesOfferListdataGridView.Sort(this.CoursesOfferListdataGridView.Columns["Course_Code"], ListSortDirection.Ascending);
                    CoursesOfferListdataGridView.CellMouseClick += new DataGridViewCellMouseEventHandler(CourseNameClick);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
                }
            }
            conn.Close();
        }

        private void CoursesOfferReportTap_Click(object sender, EventArgs e)
        {
              
        }

        private void Managebutton_Click(object sender, EventArgs e)
        {
            
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            String value = StudentSearchReporttextBox.Text.ToString();
            this.dataTable1TableAdapter.FillBy(this.dataSet1.DataTable1, value);
            this.reportViewer1.RefreshReport();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void courseManageButton_Click(object sender, EventArgs e)
        {
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            string course_code = CoursesOfferListdataGridView.CurrentRow.Cells[0].FormattedValue.ToString();
            string course_name = CoursesOfferListdataGridView.CurrentRow.Cells[1].FormattedValue.ToString();
            ((Control)this.CourseManageTap).Enabled = true;
            CoursesOfferTabControl.SelectTab(CourseManageTap);
            label25.Text = (course_code);
            label30.Text = course_name;
            String strSql;
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";

            try
            {
                conn.Open();
                strSql = "SELECT COUNT (*) FROM Course_Offer WHERE Course_Code = '"+course_code+"'";
                OleDbCommand cmdCount = new OleDbCommand(strSql, conn);
                OleDbDataReader readCount = cmdCount.ExecuteReader();
                while (readCount.Read())
                {
                    int count = readCount.GetInt32(0);
                    if (count > 0)
                    {
                        strSql = "SELECT Course_Offer.Course_Code, Course.Course_Name, TimeTable.Course_Date, TimeTable.Course_Time, Lecturer.Lecturer_Name FROM  (((Course_Offer INNER JOIN Course ON Course_Offer.Course_Code = Course.Course_Code) INNER JOIN TimeTable ON TimeTable.Offer_ID = Course_Offer.Offer_ID) INNER JOIN Lecturer ON Lecturer.Lecturer_ID = TimeTable.Lecturer_ID) WHERE Course_Offer.Course_Code = '" + course_code + "'";
                        OleDbCommand cmd = new OleDbCommand(strSql, conn);
                        OleDbDataReader readData = cmd.ExecuteReader();
                        while (readData.Read())
                        {
                            comboBox1.Text = readData.GetString(2);
                            comboBox2.Text = readData.GetString(3);
                            comboBox3.Text = readData.GetString(4);
                            comboBox1.Enabled = false;
                            comboBox2.Enabled = false;
                            comboBox3.Enabled = false;
                            LecturerAddbutton.Text = "Edit";
                            LecturerAddbutton.Click += new EventHandler(EditCourseLecturer);
                        }
                    }
                    else
                    {
                        comboBox1.Text = "";
                        comboBox2.Text = "";
                        comboBox3.Text = "";
                        LecturerAddbutton.Text = "Add";
                        LecturerAddbutton.Click += new EventHandler(LecturerAddbutton_Click);
                    }
                }               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed due to " + "\"" + ex.Message + "\"");
            }
        }
        private void EditCourseLecturer(object sender, EventArgs e)
        {
            LecturerAddbutton.Text = "Save";
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            
            string lecturerName = comboBox3.Text;
            string courseName = label25.Text;
            string course_code = label30.Text;
            string day = comboBox1.Text;
            string time = comboBox2.Text;
            int lecturerID, offerID;

            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + appPath + "\\CourseManagementDatabase.accdb";
            try
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT Lecturer_ID FROM Lecturer WHERE Lecturer_Name = '" + lecturerName + "'", conn);
                OleDbDataReader readLecturerID = cmd.ExecuteReader();
                while (readLecturerID.Read())
                {
                    lecturerID = readLecturerID.GetInt32(0);
                    OleDbCommand cmd2 = new OleDbCommand("UPDATE Course_Offer SET [Lecturer_ID] = " + lecturerID + " WHERE [Course_Code] = '" + courseName + "'", conn);
                    cmd2.ExecuteNonQuery();

                    OleDbCommand cmdTimeTable = new OleDbCommand("SELECT Offer_ID FROM Course_Offer WHERE Course_Code = '" + courseName + "'", conn);
                    OleDbDataReader readOfferID = cmdTimeTable.ExecuteReader();
                    while (readOfferID.Read())
                    {
                        offerID = readOfferID.GetInt32(0);
                        OleDbCommand cmd3 = new OleDbCommand("UPDATE Timetable SET [Offer_ID] = " + offerID + ", [Course_Date] = '" + day + "', [Course_Time] = '" + time + "', [Lecturer_ID] = " + lecturerID + " WHERE [Course_Code] = '" + course_code + "'", conn);
                        cmd3.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                LecturerAddbutton.Text = "Edit";
            }
        }
        private void CoursesOfferListdataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}