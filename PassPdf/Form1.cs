using NPOI.HPSF;
using OfficeOpenXml;
using System;
using System.Data;
using System.Windows.Forms;

namespace PassPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|Excel Macro (*.xlsm)|*.xlsm";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Title = "Select an Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFile.Text = openFileDialog.FileName;
                }
            }
        }

        private void btnBrowseExport_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select Export Folder";
                folderDialog.UseDescriptionForTitle = true; // This shows the description as the dialog title
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExportFolder.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                var excelManager = new ExcelManager(txtFile.Text);
                var employeeNames = excelManager.GetEmployeeNames();
                txtResult.Text = string.Join(Environment.NewLine, employeeNames);

                foreach ( var employee in employeeNames )
                {
                    // Populate the combo box with employee names
                    excelManager.FillEmployeeName(employee);
                    var empName = excelManager.ConvertVietnameseName(employee);
                    string pdfPath = Path.Combine(txtExportFolder.Text, $"{empName }.pdf");
                    excelManager.PrintExcelSheetToPdf(txtFile.Text, pdfPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSetPass_Click(object sender, EventArgs e)
        {

        }
    }
}
