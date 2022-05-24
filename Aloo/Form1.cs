
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace Aloo
{
    public partial class Form1 : Form
    {
      public Form1()
        {
            InitializeComponent();
        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //To where your opendialog box get starting location. My initial directory location is desktop.
            openFileDialog1.InitialDirectory = "C://Desktop";
            //Your opendialog box title name.
            openFileDialog1.Title = "Select file to be upload.";
            //which type file format you want to upload in database. just add them.
            openFileDialog1.Filter = "Select Valid Document(*.pdf; *.doc; *.xlsx; *.html)|*.pdf; *.docx; *.xlsx; *.html";
            //FilterIndex property represents the index of the filter currently selected in the file dialog box.
            openFileDialog1.FilterIndex = 1;
            try
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog1.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(openFileDialog1.FileName);
                        label1.Text = path;
                    }
                }
                else
                {
                    MessageBox.Show("Please Upload document.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult results = folderBrowserDialog1.ShowDialog();
            if (results == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog1.SelectedPath))
            {
                string[] files = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
                label3.Text = folderBrowserDialog1.SelectedPath;
            }

    

        }

        public class excelFile
        {
            public int PatientId { get; set; }
            public string FirstName  { get; set; }
            public string LastName { get; set; }
            public DateTime DateOfBirth{ get; set; }
            public int IsActive{ get; set; }

            public string FileName { get; set; }

        }


        private void Save_Click(object sender, EventArgs e)
        {
            string pdfFile = folderBrowserDialog1.SelectedPath;

            string[] pdfFileArray = Directory.GetFiles(pdfFile);



            List<excelFile> excelFiles = new List<excelFile>();




            string path = System.IO.Path.GetFullPath(openFileDialog1.FileName);
            System.Data.DataTable dtable = ConvertExcelToDataTable(path);



            if (dtable.Rows.Count > 0)
            {
                foreach (DataRow dr in dtable.Rows)
                {
                    if (dtable.Columns.Contains("PatientId"))
                    {
                        excelFile excelFile = new excelFile();

                        string F1 = dr["PatientId"].ToString();
                        excelFile.PatientId = Convert.ToInt32(F1);

                        string FirstName = dr["FirstName"].ToString();
                        excelFile.FirstName = FirstName;

                        string LastName = dr["LastName"].ToString();
                        excelFile.LastName = Regex.Replace(LastName, "[^a-zA-Z0-9_]+", ""); 

                        string DateOfBirth = dr["DateOfBirth"].ToString();
                        DateTime dateTime10 = Convert.ToDateTime(DateOfBirth);
                        excelFile.DateOfBirth = dateTime10;

                        string IsActive = dr["IsActive"].ToString();
                        excelFile.IsActive = Convert.ToInt32(IsActive);

                        var dto = dateTime10.ToString("MM-dd-yyyy");

                        var fileName = excelFile.LastName + dto;
                        fileName= fileName.Replace(" ", "");
                        fileName = Regex.Replace(fileName, "[^a-zA-Z0-9_]+", "");
                        fileName = fileName.ToLower();


                        excelFile.FileName = fileName;

                        excelFiles.Add(excelFile);
                    }
                }
            }

            foreach (var item in pdfFileArray)
            {
                var oldFilePath = item;
                var oldFileName =  Path.GetFileNameWithoutExtension(oldFilePath);

                var oldFileNameWOSpace = oldFileName.Replace(" ", "");

                oldFileNameWOSpace = oldFileNameWOSpace.Remove(0, 1);
                string oldFileNameArray = Regex.Replace(oldFileNameWOSpace, "[^a-zA-Z0-9_]+", "");
                oldFileNameArray = oldFileNameArray.Replace("_", "");
                oldFileNameArray = oldFileNameArray.ToLower();


                

                   

                    var excelRecord = excelFiles.Where(x => x.FileName == oldFileNameArray ).FirstOrDefault();
                    if (excelRecord != null)
                    {
                        var directoryPath = Path.GetDirectoryName(oldFilePath);

                        var dateTimeDir = DateTime.Now.ToString("dd-MM-yyyy");
                        var dateFolder = Path.Combine(directoryPath, dateTimeDir);

                        if (!Directory.Exists(dateFolder))
                        {
                            Directory.CreateDirectory(dateFolder);
                        }

                        var newPathName= Path.Combine(dateFolder, excelRecord.PatientId.ToString()+".pdf");

                        System.IO.File.Move(oldFilePath, newPathName);

                    AddText(path, excelRecord.PatientId);

                    }



                





            }



            MessageBox.Show("Done");  




            }



       


        public static System.Data.DataTable ConvertExcelToDataTable(string FileName)
        {
            System.Data.DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                System.Data.DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }


        public void AddText(string FileName,int patientId)
        {

            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";" +
                       @"Extended Properties='Excel 12.0;HDR=Yes;';Persist Security Info=False;";

            int totalSheet = 0;
            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();
                try
                {
                    System.Data.DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = string.Empty;
                    if (dt != null)
                    {
                        var tempDataTable = (from dataRow in dt.AsEnumerable()
                                             where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                             select dataRow).CopyToDataTable();
                        dt = tempDataTable;
                        totalSheet = dt.Rows.Count;
                        sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    }


                    OleDbCommand cmd = new OleDbCommand("UPDATE  [" + sheetName + "] SET Status= 'Added' where PatientId=" + patientId , connection);

                    cmd.ExecuteNonQuery();
                    connection.Close();


                }
                catch (Exception ex) { }
            }

        }
      
    }
}
