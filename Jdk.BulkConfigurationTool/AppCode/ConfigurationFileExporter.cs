using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System;
using Jdk.BulkConfigurationTool.Helpers;
using System.Linq;
using System.Runtime.InteropServices;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal class ConfigurationFileExporter
    {
        private readonly ConfigurationFile definition = new ConfigurationFile();
        private readonly string filePath;

        public ConfigurationFileExporter(string filePath)
        {
            this.filePath = filePath;
        }

        public event EventHandler<string> RaiseError;
        public event EventHandler<string> RaiseSuccess;

        public void Export()
        {
            try
            {
                GenerateWorkbook();

                OnRaiseSuccess($"Template exported to {filePath}");
            }
            catch (Exception e)
            {
                OnRaiseError(e.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected virtual void OnRaiseError(string e)
        {
            RaiseError?.Invoke(this, e);
        }

        protected virtual void OnRaiseSuccess(string e)
        {
            RaiseSuccess?.Invoke(this, e);
        }

        private void GenerateSheet(Excel.Worksheet sheet, string sheetName)
        {
            sheet.Name = sheetName;
            sheet.Cells.Locked = false;
            var columns = definition.Worksheets[EnumUtils.EnumValue<ConfigurationFile.WorkSheets>(sheetName)].Columns;
            foreach (var column in columns)
            {

                if (EnumUtils.IsOptional(column.TargetField))
                {
                    sheet.Cells[1, column.Position] = "(Optional)";
                    sheet.Columns[column.Position].Style = "Optional";
                }
                else
                {

                    sheet.Cells[1, column.Position] = "(Required)";
                }
                sheet.Cells[1, column.Position].Style = "HeadingDetail";

                sheet.Cells[2, column.Position] = column.Heading;
                sheet.Cells[2, column.Position].Style = "HeadingTitle";

                var defaultValue = EnumUtils.GetDefault(column.TargetField);
                if (defaultValue != null)
                {
                    sheet.Cells[3, column.Position] = $"(Default: {defaultValue.ToString()})";
                }
                sheet.Cells[3, column.Position].Style = "HeadingDetail";

                var options = EnumUtils.GetOptions(column.TargetField);
                if (options.Count() > 0)
                {
                    var cell = (Excel.Range)sheet.Columns[column.Position];
                    cell.Validation.Add(
                        Excel.XlDVType.xlValidateList,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlBetween,
                        string.Join(", ", options));
                }
            }
            sheet.Columns.AutoFit();

            sheet.Protect(Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, true,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);

            // Freeze the header rows.
            sheet.Activate();
            sheet.Application.ActiveWindow.SplitRow = 3;
            sheet.Application.ActiveWindow.FreezePanes = true;

        }

        private void GenerateWorkbook()
        {
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkBook = null;

            try
            {
                excelApp = new Excel.Application();
                excelWorkBook = excelApp.Workbooks.Add();

                var headingStyle = excelWorkBook.Styles.Add("HeadingTitle");
                headingStyle.Font.Bold = true;
                headingStyle.Font.Size = 11;
                headingStyle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headingStyle.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                headingStyle.Locked = true;

                var detailStyle = excelWorkBook.Styles.Add("HeadingDetail");
                detailStyle.Font.Bold = false;
                detailStyle.Font.Size = 9;
                detailStyle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                detailStyle.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                detailStyle.Locked = true;

                var optionalStyle = excelWorkBook.Styles.Add("Optional");
                optionalStyle.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                optionalStyle.Locked = false;

                var entitiesSheet = (Excel.Worksheet)excelWorkBook.ActiveSheet;
                GenerateSheet(entitiesSheet,
                    EnumUtils.Label(ConfigurationFile.WorkSheets.Entities));
                var attributesSheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add(After: entitiesSheet);
                GenerateSheet(attributesSheet,
                    EnumUtils.Label(ConfigurationFile.WorkSheets.Attributes));
                var oneToManySheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add(After: attributesSheet);
                GenerateSheet(oneToManySheet,
                    EnumUtils.Label(ConfigurationFile.WorkSheets.OneToManyRelationships));
                var manyToManySheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add(After: oneToManySheet);
                GenerateSheet(manyToManySheet,
                    EnumUtils.Label(ConfigurationFile.WorkSheets.ManyToManyRelationships));
                var optionSetSheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add(After: manyToManySheet);
                GenerateSheet(optionSetSheet,
                    EnumUtils.Label(ConfigurationFile.WorkSheets.OptionSets));

                entitiesSheet.Activate();
                excelWorkBook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
                    Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlLocalSessionChanges, true,
                    Missing.Value, Missing.Value, Missing.Value);

                Marshal.FinalReleaseComObject(entitiesSheet);
                Marshal.FinalReleaseComObject(attributesSheet);
                Marshal.FinalReleaseComObject(oneToManySheet);
                Marshal.FinalReleaseComObject(manyToManySheet);
                Marshal.FinalReleaseComObject(optionSetSheet);

            }
            finally
            {
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
                    Marshal.FinalReleaseComObject(excelWorkBook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.FinalReleaseComObject(excelApp);
                }
            }
        }

    }
}
