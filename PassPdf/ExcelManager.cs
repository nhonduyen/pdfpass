using System.Data;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Path = System.IO.Path;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using IWorkbook = NPOI.SS.UserModel.IWorkbook;
using System.Globalization;
using iText.Kernel.Pdf;

namespace PassPdf
{
    public class ExcelManager
    {
        private readonly string filePath;

        public ExcelManager(string filePath)
        {
            this.filePath = filePath;
        }

        public async Task FillEmployeeNameAsync(string employeeName)
        {
            try
            {
                IWorkbook workbook;
                await using (var inputStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
                {
                    if (Path.GetExtension(filePath).ToLower() == ".xls")
                        workbook = new HSSFWorkbook(inputStream);
                    else
                        workbook = new XSSFWorkbook(inputStream);
                }

                ISheet sheet = workbook.GetSheet("Payslip-Gross");
                ISheet sheetPayroll = workbook.GetSheet("Payroll");
                if (sheet == null)
                    throw new Exception("Payslip-Gross sheet not found");

                IRow row = sheet.GetRow(6) ?? sheet.CreateRow(6);
                ICell cell = row.GetCell(1) ?? row.CreateCell(1);

                ICellStyle existingStyle = cell.CellStyle;
                cell.SetCellValue(employeeName.Trim());

                if (existingStyle != null)
                    cell.CellStyle = existingStyle;

                sheet.ForceFormulaRecalculation = true;
                sheetPayroll.ForceFormulaRecalculation = true;

                await using (var outputStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
                {
                    workbook.Write(outputStream);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error filling employee name: {ex.Message}");
            }
        }

        public string ConvertVietnameseName(string name)
        {
            return name.Normalize(NormalizationForm.FormD)
                       .Replace("\u0111", "d")
                       .Replace("\u0110", "D")
                       .Where(c => !CharUnicodeInfo.GetUnicodeCategory(c).Equals(UnicodeCategory.NonSpacingMark))
                       .Aggregate("", (current, c) => current + c);
        }

        public async Task<List<Employee>> GetEmployeesAsync()
        {
            try
            {
                await using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true))
                {
                    IWorkbook workbook;
                    if (Path.GetExtension(filePath).ToLower() == ".xls")
                        workbook = new HSSFWorkbook(stream);
                    else
                        workbook = new XSSFWorkbook(stream);

                    ISheet sheet = workbook.GetSheet("Payroll");
                    if (sheet == null)
                        throw new Exception("Payroll sheet not found");

                    var employees = new List<Employee>();
                    int currentRow = 11;

                    while (true)
                    {
                        IRow row = sheet.GetRow(currentRow);
                        if (row == null) break;

                        ICell cellName = row.GetCell(1);
                        ICell cellPw = row.GetCell(41);

                        string name = cellName.StringCellValue.Trim();
                        if (name.Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                            break;

                        string cellPassword = cellPw.StringCellValue.Trim();
                        var employee = new Employee(name, ConvertVietnameseName(name), cellPassword);
                        employees.Add(employee);

                        currentRow++;
                    }

                    return employees;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading Payroll data: {ex.Message}");
            }
        }

        public async Task PrintExcelSheetToPdfAsync(string excelPath, string outputPdfPath)
        {
            string sheetName = "Payslip-Gross";

            await Task.Run(() =>
            {
                Excel.Application excelApp = null;
                Excel.Workbook workbook = null;

                try
                {
                    excelApp = new Excel.Application
                    {
                        Visible = false,
                        ScreenUpdating = false,
                        DisplayAlerts = false
                    };

                    workbook = excelApp.Workbooks.Open(excelPath, ReadOnly: true);

                    Excel.Worksheet worksheet = workbook.Sheets.Cast<Excel.Worksheet>()
                        .FirstOrDefault(ws => ws.Name == sheetName);

                    if (worksheet == null)
                        throw new Exception($"Sheet '{sheetName}' not found.");

                    foreach (Excel.Worksheet ws in workbook.Sheets)
                    {
                        ws.Visible = (ws == worksheet)
                            ? Excel.XlSheetVisibility.xlSheetVisible
                            : Excel.XlSheetVisibility.xlSheetVeryHidden;
                    }

                    worksheet.Select(Type.Missing);

                    workbook.ExportAsFixedFormat(
                        Excel.XlFixedFormatType.xlTypePDF,
                        outputPdfPath,
                        Excel.XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        From: Type.Missing,
                        To: Type.Missing,
                        OpenAfterPublish: false,
                        Type.Missing
                    );

                    if (!File.Exists(outputPdfPath))
                        throw new Exception("Export failed. PDF file was not created.");
                }
                finally
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
        }

        public async Task ProtectPdfWithPasswordAsync(string pdfPath, string userPassword, string ownerPassword)
        {
            string tempPath = Path.Combine(
                Path.GetDirectoryName(pdfPath),
                Guid.NewGuid() + ".pdf"
            );

            try
            {
                await Task.Run(() =>
                {
                    WriterProperties writerProperties = new WriterProperties();
                    writerProperties.SetStandardEncryption(
                        Encoding.UTF8.GetBytes(userPassword),
                        Encoding.UTF8.GetBytes(ownerPassword),
                        EncryptionConstants.ALLOW_PRINTING | EncryptionConstants.ALLOW_SCREENREADERS,
                        EncryptionConstants.ENCRYPTION_AES_128 | EncryptionConstants.DO_NOT_ENCRYPT_METADATA
                    );

                    using (var reader = new PdfReader(pdfPath))
                    using (var writer = new PdfWriter(tempPath, writerProperties))
                    using (var pdfDoc = new PdfDocument(reader, writer))
                    {
                        pdfDoc.Close();
                    }

                    if (File.Exists(pdfPath))
                        File.Replace(tempPath, pdfPath, null);
                    else
                        File.Move(tempPath, pdfPath);
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error protecting PDF: {ex.Message}", ex);
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }
    }
}
