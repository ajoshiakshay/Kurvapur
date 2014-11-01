using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;
using KurvapurSevaApp.DBConn;
using KurvapurSevaApp.Utilities;
using System.Drawing.Text;

namespace KurvapurSevaApp
{
    public partial class KurvapurForm : Form
    {
        private DBConnection dbConn = new DBConnection();
        private SqlConnection sqlConn;
        private string strSevaReceipt = string.Empty;
        private Logger log = new Logger();
        string strPrintData = string.Empty;
        Font fontPrint = new Font("Arial", 12, FontStyle.Bold);

        public KurvapurForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cbRashi.SelectedIndex = 0;
            cbOffRashi.SelectedIndex = 0;
            cbSearchRashi.SelectedIndex = 0;
            upcomingEvents();
        }

        private void pbCalendar_Click(object sender, EventArgs e)
        {
            DateTimePicker dtDateOfBirth = new DateTimePicker();
            dtDateOfBirth.Show();
        }

        private void dtpDOBSearch_ValueChanged(object sender, EventArgs e)
        {
            tbSearchCountry.Text = dtpSearchDOB.Value.ToString("dd-MM-yyyy");
        }

        private void dtpDateOfBirth_ValueChanged(object sender, EventArgs e)
        {
            tbDateOfBith.Text = dtpDateOfBirth.Value.ToString("dd-MM-yyyy");
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            DateTime dtDateOfBirth = new DateTime();
            string sRashi = string.Empty;
            string insertString = string.Empty;
            sqlConn = dbConn.getSQLConn();

            try
            {
                log.Log("DEBUG", "ENTERING btnSave_Click method");

                if (string.IsNullOrEmpty(tbFirstName.Text))
                {
                    MessageBox.Show("Please enter First Name", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (!tbDateOfBith.Text.Trim().Equals(string.Empty))
                {
                    if (!DateTime.TryParse(tbDateOfBith.Text.Trim(), out dtDateOfBirth))
                    {
                        MessageBox.Show("Please enter valid Date Of Birth", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                if (cbRashi.SelectedIndex == 0)
                {
                    sRashi = string.Empty;
                }
                else
                {
                    sRashi = cbRashi.Text;
                }

                //Open the connection
                dbConn.openSQLConn();

                if (tbDateOfBith.Text.Trim().Equals(string.Empty))
                {
                    insertString = "insert into Bhakt (FirstName, LastName, FatherName, MotherName," +
                                    "PhoneNumber, Address1, Address2, City, Pincode, StateRegion," +
                                    "Country, Gotra, Nakshatra, Rashi, Comments)" + 
                                    "values (@FirstName, @LastName, @FatherName, @MotherName," + 
                                    "@PhoneNumber, @Address1, @Address2, @City, @Pincode, @StateRegion," +
                                    "@Country, @Gotra, @Nakshatra, @Rashi, @Comments)";
                }
                else
                {
                    insertString = "insert into Bhakt (FirstName, LastName, FatherName, MotherName, DateOfBirth," +
                                    "PhoneNumber, Address1, Address2, City, Pincode, StateRegion," +
                                    "Country, Gotra, Nakshatra, Rashi, Comments)" +
                                    "values (@FirstName, @LastName, @FatherName, @MotherName, @DateOfBirth," +
                                    "@PhoneNumber, @Address1, @Address2, @City, @Pincode, @StateRegion," +
                                    "@Country, @Gotra, @Nakshatra, @Rashi, @Comments)";

                }

                log.Log("DEBUG", "BHAKT INSERT QUERY : " + insertString);
                
                //Pass the connection to a command object
                SqlCommand cmd = new SqlCommand(insertString, sqlConn);

                cmd.Parameters.AddWithValue("@FirstName", tbFirstName.Text);
                cmd.Parameters.AddWithValue("@LastName", tbLastName.Text);
                cmd.Parameters.AddWithValue("@FatherName", tbFatherName.Text);
                cmd.Parameters.AddWithValue("@MotherName", tbMotherName.Text);
                
                log.Log("DEBUG", "@FirstName : " + tbFirstName.Text + " @LastName : " + tbLastName.Text +
                                    " @FatherName : " + tbFatherName.Text + "@MotherName : " + tbMotherName.Text);

                if (!tbDateOfBith.Text.Trim().Equals(string.Empty))
                {
                    cmd.Parameters.AddWithValue("@DateOfBirth", dtDateOfBirth);
                    log.Log("DEBUG", "@DateOfBirth : " + dtDateOfBirth.ToString());
                }
                
                cmd.Parameters.AddWithValue("@PhoneNumber", tbPhoneNumber.Text);
                cmd.Parameters.AddWithValue("@Address1", tbAddressLine1.Text);
                cmd.Parameters.AddWithValue("@Address2", tbAddressLine2.Text);
                cmd.Parameters.AddWithValue("@City", tbCity.Text);
                cmd.Parameters.AddWithValue("@Pincode", tbPincode.Text);
                cmd.Parameters.AddWithValue("@StateRegion", tbState.Text);
                cmd.Parameters.AddWithValue("@Country", tbCountry.Text);
                cmd.Parameters.AddWithValue("@Gotra", tbGotra.Text);
                cmd.Parameters.AddWithValue("@Nakshatra", tbNakshatra.Text);
                cmd.Parameters.AddWithValue("@Rashi", sRashi);
                cmd.Parameters.AddWithValue("@Comments", tbComments.Text);

                log.Log("DEBUG", "@PhoneNumber : " + tbPhoneNumber.Text + " @Address1 :" + tbAddressLine1.Text +
                                    " @Address2 : " + tbAddressLine2.Text + " @City : " + tbCity.Text +
                                    " @City : " + tbCity.Text + " @Pincode : " + tbPincode.Text +
                                    " @StateRegion : " + tbState.Text + " @Country : " + tbCountry.Text +
                                    " @Gotra : " + tbGotra.Text + " @Nakshatra : " + tbNakshatra.Text +
                                    " @Rashi : " + sRashi + " @Comments : " + tbComments.Text);

                //get query results
                int iRowsInserted = cmd.ExecuteNonQuery();

                if(iRowsInserted >0)
                    MessageBox.Show("Data Saved Successfully", "INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show("Failed to Save the Data ", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                ResetBhaktDetails();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : Failed to Save the Data", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                log.Log("ERROR", "EXCEPTION in btnSave_Click method : " + ex.Message);
            }
            finally
            {
                //Close the connection
                if (sqlConn != null)
                {
                    dbConn.closeSQLConn();
                    log.Log("DEBUG", "CLOSE SQL CONN in btnSave_Click method");
                }
                log.Log("DEBUG", "LEAVING btnSave_Click method");
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            ResetBhaktDetails();
        }

        private void ResetBhaktDetails()
        {
            tbFirstName.Text = string.Empty;
            tbLastName.Text = string.Empty;
            tbCity.Text = string.Empty;
            tbPincode.Text = string.Empty;
            tbAddressLine2.Text = string.Empty;
            tbFatherName.Text = string.Empty;
            tbMotherName.Text = string.Empty;
            tbDateOfBith.Text = string.Empty;
            tbCountry.Text = string.Empty;
            tbPhoneNumber.Text = string.Empty;
            tbAddressLine1.Text = string.Empty;
            tbGotra.Text = string.Empty;
            tbNakshatra.Text = string.Empty;
            cbRashi.SelectedIndex = 0;
            tbComments.Text = string.Empty;
            tbState.Text = string.Empty;
        }

        private void btnOffSave_Click(object sender, EventArgs e)
        {
            saveOffering(false);
        }

        private void btnOffSavePrint_Click(object sender, EventArgs e)
        {
            saveOffering(true);
        }

        private void saveOffering(bool bPrintSevaReceipt)
        {
            DateTime dtSevaStartDate = new DateTime();
            DateTime dtSevaEndDate = new DateTime();
            string sRashi = string.Empty;
            string sSevaName = string.Empty;
            int iSevaAmt = 0;
            sqlConn = dbConn.getSQLConn();

            try
            {
                log.Log("DEBUG", "ENTERING saveOffering method");

                if (string.IsNullOrEmpty(tbOffFirstName.Text))
                {
                    MessageBox.Show("Please enter First Name", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (cbOffRashi.SelectedIndex == 0)
                {
                    sRashi = string.Empty;
                }
                else
                {
                    sRashi = cbOffRashi.Text;
                }
                
                if (rbOffAbhishek.Checked == true)
                    sSevaName = rbOffAbhishek.Text;
                else if (rbOffPalkiSeva.Checked == true)
                    sSevaName = rbOffPalkiSeva.Text;
                else if (rbOffOther.Checked == true)
                {
                    if (tbOffOther.Text.Trim().Equals(string.Empty))
                    {
                        MessageBox.Show("Please Enter Seva Details for Others", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                    {
                        sSevaName = tbOffOther.Text;
                    }
                }
                else
                {
                    MessageBox.Show("Please Select Seva to offer", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (tbAmt.Text.Equals(string.Empty))
                {
                    iSevaAmt = 0;
                }
                else
                {
                    iSevaAmt = Convert.ToInt32(Math.Round(Convert.ToDecimal(tbAmt.Text)));
                }

                if (tbOffStartDate.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Please Enter Seva Start Date", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                {
                    dtSevaStartDate = Convert.ToDateTime(tbOffStartDate.Text);
                }

                if (tbOffEndDate.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Please Enter Seva End Date", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                {
                    dtSevaEndDate = Convert.ToDateTime(tbOffEndDate.Text);
                }

                //Open the connection
                dbConn.openSQLConn();

                string insertString = "insert into Seva (FirstName, LastName, FatherName, MotherName," +
                                    "Gotra, Nakshatra, Rashi, SevaName, Amount, SevaStartDate, SevaEndDate)" +
                                    "values (@FirstName, @LastName, @FatherName, @MotherName, @Gotra," +
                                    "@Nakshatra, @Rashi, @SevaName, @Amount, @SevaStartDate, @SevaEndDate)";

                log.Log("DEBUG", "SEVA INSERT QUERY : " + insertString);

                //Pass the connection to a command object
                SqlCommand cmd = new SqlCommand(insertString, sqlConn);

                cmd.Parameters.AddWithValue("@FirstName", tbOffFirstName.Text);
                cmd.Parameters.AddWithValue("@LastName", tbOffLastName.Text);
                cmd.Parameters.AddWithValue("@FatherName", tbOffFatherName.Text);
                cmd.Parameters.AddWithValue("@MotherName", tbOffMotherName.Text);
                cmd.Parameters.AddWithValue("@Gotra", tbOffGotra.Text);
                cmd.Parameters.AddWithValue("@Nakshatra", tbOffNakshatra.Text);
                cmd.Parameters.AddWithValue("@Rashi", sRashi);
                cmd.Parameters.AddWithValue("@SevaName", sSevaName);
                cmd.Parameters.AddWithValue("@Amount", iSevaAmt);
                cmd.Parameters.AddWithValue("@SevaStartDate", dtSevaStartDate);
                cmd.Parameters.AddWithValue("@SevaEndDate", dtSevaEndDate);

                log.Log("DEBUG", "@FirstName : " + tbOffFirstName.Text + " @LastName : " + tbOffLastName.Text +
                                 " @FatherName : " + tbOffFatherName.Text + " @MotherName : " + tbOffMotherName.Text +
                                 " @Gotra : " + tbOffGotra.Text + " @Nakshatra : " + tbOffNakshatra.Text +
                                 " @Rashi : " + sRashi + " @SevaName : " + sSevaName +
                                 " @Amount : " + iSevaAmt + " @SevaStartDate : " + dtSevaStartDate +
                                 " @SevaEndDate : " + dtSevaEndDate);

                //get query results
                int iRowsInserted = cmd.ExecuteNonQuery();

                if (iRowsInserted > 0)
                {
                    MessageBox.Show("Data Saved Successfully", "INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (bPrintSevaReceipt)
                    {
                        createSevaPrintData();
                        pdPrintDocument.Print();
                    }
                }
                else
                {
                    MessageBox.Show("Failed to Save the Data ", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                resetSevaDetails();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : Failed to Save the Data", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Error);
                log.Log("ERROR", "EXCEPTION in saveOffering method : " + ex.Message);
            }
            finally
            {
                //Close the connection
                if (sqlConn != null)
                {
                    dbConn.closeSQLConn();
                    log.Log("DEBUG", "CLOSE SQL CONN in saveOffering method");
                }
                log.Log("DEBUG", "LEAVING saveOffering method");
            }
        }

        private void btnOffReset_Click(object sender, EventArgs e)
        {
            resetSevaDetails();
        }

        private void resetSevaDetails()
        {
            tbOffFirstName.Text = string.Empty;
            tbOffLastName.Text = string.Empty;
            tbOffFatherName.Text = string.Empty;
            tbOffMotherName.Text = string.Empty;
            cbOffRashi.SelectedIndex = 0;
            tbOffGotra.Text = string.Empty;
            tbOffNakshatra.Text = string.Empty;
            rbOffAbhishek.Checked = false;
            rbOffPalkiSeva.Checked = false;
            rbOffOther.Checked = false;
            tbOffOther.Text = string.Empty;
            tbAmt.Text = string.Empty;
            tbOffStartDate.Text = string.Empty;
            tbOffEndDate.Text = string.Empty;
        }

        private void dtpSearchDOB_ValueChanged(object sender, EventArgs e)
        {
            tbSearchDOB.Text = dtpSearchDOB.Value.ToString("dd-MM-yyyy");
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            DateTime dtDateOfBirth = new DateTime();
            SqlDataReader sqlRdr = null;
            string sRashi = string.Empty;
            string sQuery = string.Empty;

            try
            {
                log.Log("DEBUG", "ENTERING btnSearch_Click method");

                sqlConn = dbConn.getSQLConn();
                
                if (!tbSearchDOB.Text.Trim().Equals(string.Empty))
                {
                    if (!DateTime.TryParse(tbSearchDOB.Text.Trim(), out dtDateOfBirth))
                    {
                        MessageBox.Show("Please enter valid Date Of Birth", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }

                if (cbSearchRashi.SelectedIndex != 0)
                {
                    sRashi = cbSearchRashi.Text;
                }

                //Open the connection
                dbConn.openSQLConn();

                if (tbSearchDOB.Text.Trim().Equals(string.Empty))
                {
                    sQuery = "SELECT FirstName AS [First Name], LastName as [Last Name]," +
                                "FatherName as [Father Name], MotherName as [Mother Name], " +
                                "DateOfBirth as [Date Of Birth], PhoneNumber as [Phone Number]," +
                                "Gotra, Nakshatra, Rashi, Address1, Address2, City, StateRegion as [State], Country, Pincode " +
                                "from Bhakt " +
                                "where FirstName like @FirstName or  LastName like @LastName " +
                                "or PhoneNumber = @PhoneNumber or Gotra like @Gotra " +
                                "or Rashi like @Rashi or City like @City or StateRegion like @State " +
                                "or Country like @Country or Pincode like @Pincode";
                }
                else
                {
                    sQuery = "SELECT FirstName AS [First Name], LastName as [Last Name]," +
                                "FatherName as [Father Name], MotherName as [Mother Name], " +
                                "DateOfBirth as [Date Of Birth], PhoneNumber as [Phone Number]," +
                                "Gotra, Nakshatra, Rashi, Address1, Address2, City, StateRegion as [State], Country, Pincode " +
                                "from Bhakt " +
                                "where FirstName like @FirstName or LastName like @LastName " +
                                "or DateOfBirth = @DateOfBirth or PhoneNumber = @PhoneNumber " +
                                "or Gotra like @Gotra or Rashi like @Rashi " +
                                "or City like @City or StateRegion like @State or Country like @Country " +
                                "or Pincode like @Pincode";
                }

                //Pass the connection to a command object
                SqlCommand sqlComm = new SqlCommand(sQuery, sqlConn);

                sqlComm.Parameters.AddWithValue("@FirstName", tbSearchFirstName.Text);
                sqlComm.Parameters.AddWithValue("@LastName", tbSearchLastName.Text);
                
                if (!tbSearchDOB.Text.Trim().Equals(string.Empty))
                    sqlComm.Parameters.AddWithValue("@DateOfBirth", dtDateOfBirth);

                sqlComm.Parameters.AddWithValue("@PhoneNumber", tbSearchPhoneNumber.Text);
                sqlComm.Parameters.AddWithValue("@Gotra", tbSearchGotra.Text);
                sqlComm.Parameters.AddWithValue("@Rashi", sRashi);
                sqlComm.Parameters.AddWithValue("@City", tbSearchCity.Text);
                sqlComm.Parameters.AddWithValue("@State", tbSearchState.Text);
                sqlComm.Parameters.AddWithValue("@Country", tbSearchCountry.Text);
                sqlComm.Parameters.AddWithValue("@Pincode", tbSearchPincode.Text);

                log.Log("DEBUG", "BHAKT SELECT QUERY : " + sQuery);

                sqlRdr = sqlComm.ExecuteReader();
                DataSet ds = new DataSet();
                DataTable dt = new DataTable("NewSeva");
                ds.Tables.Add(dt);
                ds.Load(sqlRdr, LoadOption.PreserveChanges, ds.Tables[0]);

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("No details found. Please enter valid details", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    dgvSearchResult.DataSource = ds.Tables[0];
                    createBhaktAddressPrintData(ds.Tables[0]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Faild to Search Bhakt details", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                log.Log("ERROR", "EXCEPTION in btnSearch_Click method : " + ex.Message);
            }
            finally
            {
                //close the reader
                if (sqlRdr != null)
                {
                    sqlRdr.Close();
                    log.Log("DEBUG", "CLOSE SQL RDR in btnSearch_Click method");
                }

                //Close the connection
                if (sqlConn != null)
                {
                    dbConn.closeSQLConn();
                    log.Log("DEBUG", "CLOSE SQL CONN in btnSearch_Click method");
                }
                log.Log("DEBUG", "LEAVING btnSearch_Click method");
            }
        }

        private void dtpSevaStartDate_ValueChanged(object sender, EventArgs e)
        {
            tbSearchSevaStartDate.Text = dtpSevaStartDate.Value.ToString("dd-MM-yyyy");
        }

        private void dtpSearchSevaEndDate_ValueChanged(object sender, EventArgs e)
        {
            tbSearchSevaEndDate.Text = dtpSearchSevaEndDate.Value.ToString("dd-MM-yyyy");
        }

        private void btnSearchSeva_Click(object sender, EventArgs e)
        {
            DateTime dtStarDate = new DateTime();
            DateTime dtEndDate = new DateTime();
            SqlDataReader sqlRdr = null;
            sqlConn = dbConn.getSQLConn();

            try
            {
                log.Log("DEBUG", "ENTERING btnSearchSeva_Click method");

                if (!tbSearchSevaStartDate.Text.Trim().Equals(string.Empty))
                {
                    if (!DateTime.TryParse(tbSearchSevaStartDate.Text.Trim(), out dtStarDate))
                    {
                        MessageBox.Show("Please enter valid Seva Start Date", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please enter Seva Start Date", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (!tbSearchSevaEndDate.Text.Trim().Equals(string.Empty))
                {
                    if (!DateTime.TryParse(tbSearchSevaEndDate.Text.Trim(), out dtEndDate))
                    {
                        MessageBox.Show("Please enter valid Seva End Date", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Please enter Seva End Date", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }


                //Open the connection
                dbConn.openSQLConn();

                string sQuery = "SELECT FirstName AS [First Name], LastName as [Last Name]," +
                                "FatherName as [Father Name], MotherName as [Mother Name], Gotra, Nakshatra," +
                                "Rashi, SevaName as Seva, SevaStartDate as [Start Date] from Seva " +
                                "where SevaStartDate >= @StarDate and  SevaStartDate <= @EndDate";

                //Pass the connection to a command object
                SqlCommand sqlComm = new SqlCommand(sQuery, sqlConn);

                sqlComm.Parameters.AddWithValue("@StarDate", dtStarDate);
                sqlComm.Parameters.AddWithValue("@EndDate", dtEndDate);

                log.Log("DEBUG", "SEVA SELECT QUERY : " + sQuery + "StarDate : " +
                        dtStarDate.ToString() + "EndDate : " + dtEndDate.ToString());

                sqlRdr = sqlComm.ExecuteReader();
                DataSet ds = new DataSet();
                DataTable dt = new DataTable("SearchSeva");
                ds.Tables.Add(dt);
                ds.Load(sqlRdr, LoadOption.PreserveChanges, ds.Tables[0]);

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("No details found. Please enter valid details", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    dgvSearchSevaResult.DataSource = ds.Tables[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Faild to Search Seva details", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                log.Log("ERROR", "EXCEPTION in btnSearchSeva_Click method : " + ex.Message);
            }
            finally
            {
                //close the reader
                if (sqlRdr != null)
                {
                    sqlRdr.Close();
                    log.Log("DEBUG", "CLOSE SQL RDR in btnSearchSeva_Click method");
                }

                //Close the connection
                if (sqlConn != null)
                {
                    dbConn.closeSQLConn();
                    log.Log("DEBUG", "CLOSE SQL CONN in btnSearchSeva_Click method");
                }
                log.Log("DEBUG", "LEAVING btnSearchSeva_Click method");
            }
        }

        private void upcomingEvents()
        {
            int iStartDateRange = Convert.ToInt32(ConfigurationSettings.AppSettings["EventStartDateRange"]);
            int iEndDateRange = Convert.ToInt32(ConfigurationSettings.AppSettings["EventEndDateRange"]);
            SqlDataReader sqlRdr = null;
            DateTime dtStarDate = DateTime.Today.AddDays(-iStartDateRange);
            DateTime dtEndDate = DateTime.Today.AddDays(iEndDateRange);

            try
            {
                log.Log("DEBUG", "ENTERING upcomingEvents method");

                sqlConn = dbConn.getSQLConn();
                //Open the connection
                dbConn.openSQLConn();

                string sQuery = "SELECT FirstName AS [First Name], LastName as [Last Name]," +
                                "FatherName as [Father Name], MotherName as [Mother Name], Gotra, Nakshatra," +
                                "Rashi, SevaName as Seva, SevaStartDate as [Start Date] from Seva " +
                                "where SevaStartDate >= @StarDate and  SevaStartDate <= @EndDate";

                //Pass the connection to a command object
                SqlCommand sqlComm = new SqlCommand(sQuery, sqlConn);

                sqlComm.Parameters.AddWithValue("@StarDate", dtStarDate);
                sqlComm.Parameters.AddWithValue("@EndDate", dtEndDate);

                log.Log("DEBUG", "Upcoming Events SELECT QUERY : " + sQuery + "StarDate : " + 
                        dtStarDate.ToString() + "EndDate : " + dtEndDate.ToString());

                sqlRdr = sqlComm.ExecuteReader();
                DataSet ds = new DataSet();
                DataTable dt = new DataTable("NewSeva");
                ds.Tables.Add(dt);
                ds.Load(sqlRdr, LoadOption.PreserveChanges, ds.Tables[0]);
                dgvUpcomingEvents.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Faild to retrieve upcoming Seva details", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                log.Log("ERROR", "EXCEPTION in upcomingEvents method : " + ex.Message);
            }
            finally
            {
                //close the reader
                if (sqlRdr != null)
                {
                    sqlRdr.Close();
                    log.Log("DEBUG", "CLOSE SQL RDR in upcomingEvents method");
                }

                //Close the connection
                if (sqlConn != null)
                {
                    dbConn.closeSQLConn();
                    log.Log("DEBUG", "CLOSE SQL CONN in upcomingEvents method");
                }
                log.Log("DEBUG", "LEAVING upcomingEvents method");
            }
        }

        private void dtpOfferingStartDate_ValueChanged(object sender, EventArgs e)
        {
            tbOffStartDate.Text = dtpOfferingStartDate.Value.ToString("dd-MM-yyyy");
        }

        private void dtpOfferingEndDate_ValueChanged(object sender, EventArgs e)
        {
            tbOffEndDate.Text = dtpOfferingEndDate.Value.ToString("dd-MM-yyyy");
        }

        private void btnFetchUpcomingSeva_Click(object sender, EventArgs e)
        {
            upcomingEvents();
        }

        private void btnBhaktPrint_Click(object sender, EventArgs e)
        {
            log.Log("DEBUG", "ENTERING btnBhaktPrint_Click method");

            pdPrintDocument.Print();

            log.Log("DEBUG", "LEAVING btnBhaktPrint_Click method");
        }

        private void pdPrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            log.Log("DEBUG", "ENTERING pdPrintDocument_PrintPage method");

            e.Graphics.DrawString(strPrintData, fontPrint, Brushes.Black, 100, 20);

            log.Log("DEBUG", "DATA for PRINT : " + strPrintData);

            log.Log("DEBUG", "LEAVING pdPrintDocument_PrintPage method");
        }

        private void btnSevaPrintPreview_Click(object sender, EventArgs e)
        {
            log.Log("DEBUG", "ENTERING btnSevaPrintPreview_Click method");

            createSevaPrintData();
            ppdPrintPreviewDialog.Document = pdPrintDocument;
            ppdPrintPreviewDialog.ShowDialog();

            log.Log("DEBUG", "LEAVING btnSevaPrintPreview_Click method");
        }

        private void createSevaPrintData()
        {
            strPrintData = string.Empty;
            string sSevaName = string.Empty;

            log.Log("DEBUG", "ENTERING createSevaPrintData method");


            if (rbOffAbhishek.Checked == true)
            {
                sSevaName = rbOffAbhishek.Text;
            }
            else if (rbOffPalkiSeva.Checked == true)
            {
                sSevaName = rbOffPalkiSeva.Text;
            }
            else if (rbOffOther.Checked == true)
            {
                sSevaName = tbOffOther.Text;
            }

            strPrintData = "                                                    SHRI KSHETRA KURVAPUR" +
             " \n\nName : " + tbOffFirstName.Text + " " + tbOffLastName.Text +
             " \n\nFather Name : " + tbOffFatherName.Text +
             " \n\nMother Name : " + tbOffMotherName.Text +
             " \n\nGotra : " + tbOffGotra.Text +
             " \n\nNakshatra : " + tbOffNakshatra.Text +
             " \n\nRashi : " + cbOffRashi.Text +
             " \n\nSeva Name : " + sSevaName +
             " \n\nAmount : " + tbAmt.Text +
             " \n\nSeva Start Date : " + tbOffStartDate.Text +
             " \n\nSeva End Date : " + tbOffEndDate.Text;

            log.Log("DEBUG", "DATA for PRINT : " + strPrintData);

            log.Log("DEBUG", "LEAVING createSevaPrintData method");
        }

        private void btnBhaktPrintPreview_Click(object sender, EventArgs e)
        {
            log.Log("DEBUG", "ENTERING btnBhaktPrintPreview_Click method");

            ppdPrintPreviewDialog.Document = pdPrintDocument;
            ppdPrintPreviewDialog.ShowDialog();

            log.Log("DEBUG", "LEAVING btnBhaktPrintPreview_Click method");
        }

        private void createBhaktAddressPrintData(DataTable dtBhaktAddressData)
        {
            strPrintData = string.Empty;

            log.Log("DEBUG", "ENTERING createBhaktAddressPrintData method");

            foreach (DataRow row in dtBhaktAddressData.Rows)
            {
                strPrintData += row["First Name"] + " " + row["Father Name"] + " " + row["Last Name"] +
                                "\n" + row["Address1"] +
                                "\n" + row["Address2"] +
                                "\n" + row["City"] + " - " + row["Pincode"] +
                                "\n" + row["State"] + 
                                "\n" + row["Country"] +
                                "\n" + "Phone Number" + " - " + row["Phone Number"] +
                                "\n\n               ------------------------------------------------                \n\n";
            }

            log.Log("DEBUG", "DATA for PRINT : " + strPrintData);

            log.Log("DEBUG", "LEAVING createBhaktAddressPrintData method");
            
        }
    }
}