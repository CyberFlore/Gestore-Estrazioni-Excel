using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ExcelTemplateManager
{
    public class ExcelManager
    {
        public ExcelManager()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public List<ColumnInfo> LoadExcelFile(string filePath)
        {
            List<ColumnInfo> columns = new List<ColumnInfo>();

            if (Path.GetExtension(filePath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception("Il formato .xls non è supportato. Si prega di salvare il file in formato .xlsx");
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int headerRow = FindHeaderRow(worksheet);

                if (headerRow == -1)
                {
                    throw new Exception("Impossibile trovare la riga delle intestazioni nel file Excel.");
                }

                int columnCount = worksheet.Dimension.End.Column;

                for (int i = 1; i <= columnCount; i++)
                {
                    var headerText = worksheet.Cells[headerRow, i].Text;
                    if (!string.IsNullOrWhiteSpace(headerText))
                    {
                        columns.Add(new ColumnInfo
                        {
                            Header = headerText,
                            OriginalIndex = i - 1
                        });
                    }
                }
            }

            return columns;
        }

        private int FindHeaderRow(ExcelWorksheet worksheet)
        {
            int maxRows = Math.Min(30, worksheet.Dimension.End.Row);

            for (int row = 1; row <= maxRows; row++)
            {
                int nonEmptyCells = 0;
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text))
                    {
                        nonEmptyCells++;
                    }
                }

                if (nonEmptyCells >= 3)
                {
                    return row;
                }
            }

            return -1;
        }

        public void ExportExcelFile(string sourceFilePath, string targetFilePath, List<ColumnInfo> columns)
        {
            if (Path.GetExtension(sourceFilePath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                throw new Exception("Il formato .xls non è supportato. Si prega di salvare il file in formato .xlsx");
            }

            using (var sourcePackage = new ExcelPackage(new FileInfo(sourceFilePath)))
            using (var targetPackage = new ExcelPackage())
            {
                var sourceWorksheet = sourcePackage.Workbook.Worksheets[0];
                var targetWorksheet = targetPackage.Workbook.Worksheets.Add("Sheet1");

                int headerRow = FindHeaderRow(sourceWorksheet);
                if (headerRow == -1)
                {
                    throw new Exception("Impossibile trovare la riga delle intestazioni nel file Excel.");
                }

                int lastRow = sourceWorksheet.Dimension.End.Row;
                int currentTargetColumn = 1;

                foreach (var column in columns)
                {
                    int sourceColumn = column.OriginalIndex + 1;

                    if (sourceColumn <= 0 || sourceColumn > sourceWorksheet.Dimension.End.Column)
                    {
                        // Crea una nuova colonna con solo l'intestazione
                        var cell = targetWorksheet.Cells[headerRow, currentTargetColumn];
                        cell.Value = column.Header; // Usa il nuovo nome della colonna
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(211, 211, 211)); // LightGray

                        // Aggiungi validazione se configurata
                        if (column.HasValidation && column.ValidationOptions.Count > 0)
                        {
                            var validation = targetWorksheet.DataValidations.AddListValidation(
                                targetWorksheet.Cells[headerRow + 1, currentTargetColumn, lastRow, currentTargetColumn].Address);
                            foreach (var option in column.ValidationOptions)
                            {
                                validation.Formula.Values.Add(option);
                            }
                        }
                    }
                    else
                    {
                        // Copia l'intera colonna con tutti i dati e la formattazione
                        var sourceRange = sourceWorksheet.Cells[1, sourceColumn, lastRow, sourceColumn];
                        var targetRange = targetWorksheet.Cells[1, currentTargetColumn, lastRow, currentTargetColumn];

                        // Copia valori e formattazione base
                        sourceRange.Copy(targetRange);

                        // Imposta il nuovo nome della colonna
                        var headerCell = targetWorksheet.Cells[headerRow, currentTargetColumn];
                        headerCell.Value = column.Header; // Usa il nuovo nome della colonna
                        headerCell.Style.Font.Bold = true;

                        // Aggiungi validazione se configurata
                        if (column.HasValidation && column.ValidationOptions.Count > 0)
                        {
                            var validation = targetWorksheet.DataValidations.AddListValidation(
                                targetWorksheet.Cells[headerRow + 1, currentTargetColumn, lastRow, currentTargetColumn].Address);
                            foreach (var option in column.ValidationOptions)
                            {
                                validation.Formula.Values.Add(option);
                            }
                        }

                        // Copia la larghezza della colonna
                        targetWorksheet.Column(currentTargetColumn).Width = sourceWorksheet.Column(sourceColumn).Width;
                    }

                    currentTargetColumn++;
                }

                // Applica l'autofit alle colonne e congela la prima riga
                targetWorksheet.Cells[targetWorksheet.Dimension.Address].AutoFitColumns();
                targetWorksheet.View.FreezePanes(headerRow + 1, 1);

                // Salva il file
                targetPackage.SaveAs(new FileInfo(targetFilePath));
            }
        }
    }
}