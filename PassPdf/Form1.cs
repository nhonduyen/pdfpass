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

        private async void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                // Check if file exists first
                if (!File.Exists(txtFile.Text))
                {
                    MessageBox.Show("File not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Check if file is read-only
                var fileAttributes = File.GetAttributes(txtFile.Text);
                if (fileAttributes.HasFlag(FileAttributes.ReadOnly))
                {
                    MessageBox.Show("File is read-only. Please remove read-only attribute.", "Error",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var excelManager = new ExcelManager(txtFile.Text);
                var employees = await excelManager.GetEmployeesAsync();
               
                int processedCount = 0;
                foreach (var employee in employees)
                {
                    try
                    {
                        var pdfPath = Path.Combine(txtExportFolder.Text, $"{txtPrefix.Text} - {employee.Name.Replace(" ","")}.pdf");
                        await excelManager.FillEmployeeNameAsync(employee.VietnameseName);
                        await excelManager.PrintExcelSheetToPdfAsync(txtFile.Text, pdfPath);
                        await excelManager.ProtectPdfWithPasswordAsync(pdfPath, employee.Password, employee.Password);
                        processedCount++;
                        txtResult.Text += $"{processedCount}. Export {pdfPath} success - Password: {employee.Password}{Environment.NewLine}";
                    }
                    catch (Exception employeeEx)
                    {
                        var result = MessageBox.Show(
                            $"Error processing employee '{employee}': {employeeEx.Message}\n\nDo you want to continue with the next employee?",
                            "Error",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning);

                        if (result == DialogResult.No)
                            break;
                    }
                }

                MessageBox.Show($"Export completed! Processed {processedCount} out of {employees.Count} employees.",
                               "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing Excel file: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                               "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSetPass_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtExportFolder.Text = @"D:\pdf";
            txtFile.Text = @"D:\pdf\1.xlsm";
        }
    }
}
