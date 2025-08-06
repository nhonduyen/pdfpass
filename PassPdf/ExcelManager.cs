using System.Data;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using iText.Kernel.Pdf;
using Path = System.IO.Path;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;
using IWorkbook = NPOI.SS.UserModel.IWorkbook;
using System.Globalization;

namespace PassPdf
{
    public class ExcelManager
    {
        private readonly string filePath;

        public ExcelManager(string filePath)
        {
            this.filePath = filePath;
        }

        public void FillEmployeeName(string employeeName)
        {
            try
            {
                IWorkbook workbook;
                using (var inputStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (System.IO.Path.GetExtension(filePath).ToLower() == ".xls")
                        workbook = new HSSFWorkbook(inputStream);
                    else
                        workbook = new XSSFWorkbook(inputStream);
                }

                ISheet sheet = workbook.GetSheet("Payslip-Gross");
                if (sheet == null)
                    throw new Exception("Payslip-Gross sheet not found");

                IRow row = sheet.GetRow(6);
                if (row == null)
                    row = sheet.CreateRow(6);

                ICell cell = row.GetCell(1);
                if (cell == null)
                    cell = row.CreateCell(1);

                ICellStyle existingStyle = cell.CellStyle;
                cell.SetCellValue(employeeName.Trim());

                if (existingStyle != null)
                {
                    cell.CellStyle = existingStyle;
                }

                // Try to evaluate only safe formulas
                try
                {
                    IFormulaEvaluator evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

                    // Force recalculation of the current cell and its dependents
                    evaluator.NotifySetFormula(cell);
                    evaluator.EvaluateFormulaCell(cell);

                    // Only evaluate formulas in the current sheet that don't have external references
                    EvaluateSafeFormulas(evaluator, sheet);
                    sheet.ForceFormulaRecalculation = true;
                }
                catch (Exception evalEx)
                {
                    // If evaluation fails, continue without it
                    Console.WriteLine($"Formula evaluation skipped: {evalEx.Message}");
                }

                using (var outputStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(outputStream);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error filling employee name: {ex.Message}");
            }
        }

        private void EvaluateSafeFormulas(IFormulaEvaluator evaluator, ISheet sheet)
        {
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                {
                    ICell cell = row.GetCell(cellIndex);
                    if (cell == null || cell.CellType != CellType.Formula) continue;

                    string formula = cell.CellFormula;

                    // Skip formulas with external references
                    if (formula.Contains("[") && formula.Contains("]"))
                    {
                        continue; // Skip external references
                    }

                    try
                    {
                        evaluator.Evaluate(cell);
                    }
                    catch
                    {
                        // Skip problematic formulas
                        continue;
                    }
                }
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

        public List<string> GetEmployeeNames()
        {
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;
                    if (System.IO.Path.GetExtension(filePath).ToLower() == ".xls")
                        workbook = new HSSFWorkbook(stream);
                    else
                        workbook = new XSSFWorkbook(stream);

                    // Get the Payroll sheet
                    ISheet sheet = workbook.GetSheet("Payroll");
                    if (sheet == null)
                        throw new Exception("Payroll sheet not found");

                    List<string> columnData = new List<string>();
                    int currentRow = 11; // Starting from row 12 (0-based index is 11)

                    while (true)
                    {
                        IRow row = sheet.GetRow(currentRow);
                        if (row == null)
                            break;

                        // Get cell from column 2 (0-based index is 1)
                        ICell cell = row.GetCell(1);
                        string cellValue = cell?.ToString() ?? string.Empty;

                        // Check if we've reached the TOTAL row
                        if (cellValue.Trim().Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
                            break;

                        // Add non-empty values to the list
                        if (!string.IsNullOrWhiteSpace(cellValue))
                            columnData.Add(cellValue);

                        currentRow++;
                    }

                    return columnData;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading Payroll data: {ex.Message}");
            }
        }

        public DataTable ReadExcelFile(string sheetName = null)
        {
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;

                    // Check file extension to use the correct reader
                    if (Path.GetExtension(filePath).ToLower() == ".xls")
                        workbook = new HSSFWorkbook(stream);
                    else
                        workbook = new XSSFWorkbook(stream);

                    // Get the specified worksheet, or the first one if no name is provided
                    ISheet sheet = string.IsNullOrEmpty(sheetName)
                        ? workbook.GetSheetAt(0)
                        : workbook.GetSheet(sheetName);

                    if (sheet == null)
                        throw new Exception("Worksheet not found");

                    DataTable dt = new DataTable();
                    IRow headerRow = sheet.GetRow(0);
                    int colCount = headerRow.LastCellNum;

                    // Create columns from header row
                    for (int i = 0; i < colCount; i++)
                    {
                        ICell cell = headerRow.GetCell(i);
                        string columnName = cell?.ToString() ?? $"Column{i + 1}";
                        dt.Columns.Add(columnName);
                    }

                    // Read data rows
                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        DataRow dataRow = dt.NewRow();
                        for (int j = 0; j < colCount; j++)
                        {
                            ICell cell = row.GetCell(j);
                            dataRow[j] = cell?.ToString() ?? string.Empty;
                        }
                        dt.Rows.Add(dataRow);
                    }

                    return dt;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading Excel file: {ex.Message}");
            }
        }

        public void PrintExcelSheetToPdf(string excelPath, string outputPdfPath)
        {
            string sheetName = "Payslip-Gross";
            var excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(excelPath);

                // Find the target sheet
                Excel.Worksheet worksheet = workbook.Sheets.Cast<Excel.Worksheet>()
                    .FirstOrDefault(ws => ws.Name == sheetName);

                if (worksheet == null)
                    throw new Exception($"Sheet '{sheetName}' not found.");

                // Hide all sheets except the one to print
                foreach (Excel.Worksheet ws in workbook.Sheets)
                {
                    ws.Visible = (ws == worksheet) ? Excel.XlSheetVisibility.xlSheetVisible : Excel.XlSheetVisibility.xlSheetVeryHidden;
                }

                // Print that single sheet to PDF
                worksheet.Select(Type.Missing);
                workbook.ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF,
                    outputPdfPath,
                    Excel.XlFixedFormatQuality.xlQualityStandard,
                    true, true, 1, 1, false, Type.Missing
                );
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void ProtectPdfWithPassword(string pdfPath, string userPassword, string ownerPassword)
        {
            try
            {
                // Read the existing PDF
                byte[] pdfBytes = File.ReadAllBytes(pdfPath);
                string tempPath = Path.Combine(Path.GetDirectoryName(pdfPath), Guid.NewGuid().ToString() + ".pdf");

                WriterProperties writerProperties = new WriterProperties();
                writerProperties.SetStandardEncryption(
                    Encoding.UTF8.GetBytes(userPassword),        // Password to open the PDF
                    Encoding.UTF8.GetBytes(ownerPassword),       // Password for permissions
                    EncryptionConstants.ALLOW_PRINTING |         // Set allowed permissions
                    EncryptionConstants.ALLOW_SCREENREADERS,     
                    EncryptionConstants.ENCRYPTION_AES_128 |     // Use AES 128-bit encryption
                    EncryptionConstants.DO_NOT_ENCRYPT_METADATA
                );

                using (var reader = new PdfReader(new MemoryStream(pdfBytes)))
                using (var writer = new PdfWriter(tempPath, writerProperties))
                using (var pdfDoc = new PdfDocument(reader, writer))
                {
                    pdfDoc.Close();
                }

                // Replace original file with protected version
                File.Delete(pdfPath);
                File.Move(tempPath, pdfPath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error protecting PDF: {ex.Message}");
            }
        }

        public string[] GetWorksheetNames()
        {
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;

                    // Check file extension to use the correct reader
                    if (Path.GetExtension(filePath).ToLower() == ".xls")
                        workbook = new HSSFWorkbook(stream);
                    else
                        workbook = new XSSFWorkbook(stream);

                    string[] names = new string[workbook.NumberOfSheets];
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        names[i] = workbook.GetSheetName(i);
                    }

                    return names;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting worksheet names: {ex.Message}");
            }
        }
    }
}
