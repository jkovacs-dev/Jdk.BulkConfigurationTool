using Jdk.BulkConfigurationTool.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal class ConfigurationFileImporter
    {
        private readonly ConfigurationFile definition = new ConfigurationFile();
        private readonly string filePath;

        public ConfigurationFileImporter(string filePath)
        {
            this.filePath = filePath;
        }

        public event EventHandler<string> RaiseError;
        public event EventHandler<string> RaiseSuccess;

        public ConfigurationFile Import()
        {
            try
            {
                ReadExcelData();
                return definition;
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
            return null;
        }

        protected virtual void OnRaiseError(string e)
        {
            RaiseError?.Invoke(this, e);
        }

        protected virtual void OnRaiseSuccess(string e)
        {
            RaiseSuccess?.Invoke(this, e);
        }

        private static void MapColumnHeaders(Excel.Worksheet xlSheet, List<ConfigurationFile.Column> columns)
        {
            var columnCount = xlSheet.UsedRange.Columns.Count;
            columns.ForEach(x => x.Position = 0);
            for (int c = 1; c <= columnCount; c++)
            {
                if (xlSheet.Cells[2, c].Value != null)
                {
                    var sheetValue = xlSheet.Cells[2, c].Value as string;
                    var mappedColumn = columns.FirstOrDefault(x => string.Equals(x.Heading, sheetValue));
                    if (mappedColumn != null)
                    {
                        mappedColumn.Position = c;
                    }
                }
            }
            var missingMandatory = columns.Where(x => x.Position == 0 && !x.Optional);
            if (missingMandatory.Count() > 0)
            {
                throw new FormatException($"Unable to process worksheet {xlSheet.Name}. Missing mandatory columns: {string.Join(", ", missingMandatory.Select(x => x.Heading))}.");
            }
        }
        private static string RangeAddress(Excel.Range rng) => rng.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,
                   Missing.Value, Missing.Value);

        private static void ReadData(Excel.Worksheet xlSheet, ConfigurationFile.Worksheet configSheet)
        {
            MapColumnHeaders(xlSheet, configSheet.Columns);
            var range = xlSheet.UsedRange;
            var rangeAddress = RangeAddress(range);
            object[,] values = range.Value;
            for (int r = 3; r < values.GetLength(0); r++)
            {
                var data = new object[values.GetLength(1)];
                for (int c = 0; c < values.GetLength(1); c++)
                {
                    data[c] = values[r + 1, c + 1];
                }
                configSheet.Data.Add(data);
            }
        }

        private void ReadExcelData()
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(filePath, Missing.Value, true);
                var sheetEnumerator = xlWorkbook.Sheets.GetEnumerator();
                while (sheetEnumerator.MoveNext())
                {
                    var sheet = (Excel.Worksheet)sheetEnumerator.Current;
                    if (string.Equals(sheet.Name, EnumUtils.Label(ConfigurationFile.WorkSheets.Entities)))
                    {
                        // read entities sheet
                        ReadData(sheet, definition.Worksheets[ConfigurationFile.WorkSheets.Entities]);
                    }
                    if (string.Equals(sheet.Name, EnumUtils.Label(ConfigurationFile.WorkSheets.Attributes)))
                    {
                        // read attributes sheet
                        ReadData(sheet, definition.Worksheets[ConfigurationFile.WorkSheets.Attributes]);
                    }
                    if (string.Equals(sheet.Name, EnumUtils.Label(ConfigurationFile.WorkSheets.OneToManyRelationships)))
                    {
                        // read oneToMany sheet
                        ReadData(sheet, definition.Worksheets[ConfigurationFile.WorkSheets.OneToManyRelationships]);
                    }
                    if (string.Equals(sheet.Name, EnumUtils.Label(ConfigurationFile.WorkSheets.ManyToManyRelationships)))
                    {
                        // read manyToMany sheet
                        ReadData(sheet, definition.Worksheets[ConfigurationFile.WorkSheets.ManyToManyRelationships]);
                    }
                    if (string.Equals(sheet.Name, EnumUtils.Label(ConfigurationFile.WorkSheets.OptionSets)))
                    {
                        // read optionSet sheet
                        ReadData(sheet, definition.Worksheets[ConfigurationFile.WorkSheets.OptionSets]);
                    }
                }
            }
            finally
            {
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                    Marshal.FinalReleaseComObject(xlWorkbook);
                }

                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.FinalReleaseComObject(xlApp);
                }
            }
        }

    }
}
