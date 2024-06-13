using CsvHelper;
using OfficeOpenXml;
using System.Configuration;
using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using static OfficeOpenXml.ExcelErrorValue;
using System.Reflection.PortableExecutable;

class ExcelReaderWriter
{
    static void Main(string[] args)
    {
        // Set the license context to NonCommercial for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        try
        {
            // Retrieve values from App.config of file path
            string outputDirectory = System.Configuration.ConfigurationManager.AppSettings["OutputDirectory"] ?? "";
            string outputFileName = System.Configuration.ConfigurationManager.AppSettings["OutputFileName"] ?? "";

            // Hardcoded Values from App.config for columns
            Dictionary<string, string> hardcodedValues = new Dictionary<string, string>
            {
                { "Tillverkare", System.Configuration.ConfigurationManager.AppSettings["Tillverkare"] ?? "" },
                { "Extra kategori 1", System.Configuration.ConfigurationManager.AppSettings["ExtraKategori1"] ?? "" },
                { "Extra kategori 2", System.Configuration.ConfigurationManager.AppSettings["ExtraKategori2"] ?? "" },
                { "Sortering", System.Configuration.ConfigurationManager.AppSettings["Sortering"] ?? "" },
                { "Dold", System.Configuration.ConfigurationManager.AppSettings["Dold"] ?? "" },
                { "Aktiv", System.Configuration.ConfigurationManager.AppSettings["Aktiv"] ?? "" },
                { "Leverantör", System.Configuration.ConfigurationManager.AppSettings["Leverantör"] ?? "" },
                { "ShopId 278 Driva Lehmanns AB", System.Configuration.ConfigurationManager.AppSettings["ShopId278"] ?? "" },
                { "ShopId 1 fronta.se", System.Configuration.ConfigurationManager.AppSettings["ShopId1"] ?? "" }
            };

            // Group to Huvudkategori mapping 
            Dictionary<string, string> groupToHuvudkategori = new Dictionary<string, string>
            {
                { "T-SHIRTS", System.Configuration.ConfigurationManager.AppSettings["Group_T-SHIRTS"] ?? "" },
                { "POLOS", System.Configuration.ConfigurationManager.AppSettings["Group_POLOS"] ?? "" },
                { "SHIRTS", System.Configuration.ConfigurationManager.AppSettings["Group_SHIRTS"] ?? "" },
                { "SWEATS", System.Configuration.ConfigurationManager.AppSettings["Group_SWEATSHIRTS"] ?? "" },
                { "KNITWEAR", System.Configuration.ConfigurationManager.AppSettings["Group_KNITWEAR"] ?? "" },
                { "FLEECE", System.Configuration.ConfigurationManager.AppSettings["Group_FLEACE"] ?? "" },
                { "OUTERWEAR", System.Configuration.ConfigurationManager.AppSettings["Group_OUTER WEAR"] ?? "" }
            };

            // Prompt user for the second input file path until a valid path is provided
            string inputFilePath2;
            do
            {
                Console.WriteLine("Enter the file path for the Excel or CSV file (00SWEFROHUV data file):");
                inputFilePath2 = Console.ReadLine();
            } while (!File.Exists(inputFilePath2));

            // Convert CSV file to Excel format, if necessary
            inputFilePath2 = ConvertToExcelIfCSV(inputFilePath2);

            // Generate a unique file name for the output
            string outputPath = GenerateUniqueFileName(outputDirectory, outputFileName);

            // Load the Excel file
            using var package2 = new ExcelPackage(new FileInfo(inputFilePath2));

            // Get the first worksheet in the file (assuming single sheet per file)
            var worksheet2 = package2.Workbook.Worksheets[0];

            // Define the translation dictionary for headers mapping columns
            var translationDictionary = new Dictionary<string, string>
            {
                { "Namn", "Product_Name" },
                { "Artikelnr produkt", "TJ_Style_no" },
                { "Artikelnr variant", "CSV_Code" },
                { "Pris", "Grp_A_Price" },
                { "Vikt", "Weight" },
                { "Tillverkare", "Manufacturer" },
                { "Lager", "Inventory" },
                { "Huvudbildadress", "Model" },
                { "Kompleterandebildadress1", "Front" },
                { "Kompleterandebildadress2", "Back" },
                { "Kompleterandebildadress3", "Left" },
                { "Beskrivning", "Description" },
                { "Storlek", "Size" },
                { "Färg", "Colour" },                
                { "Huvudkategori", "Main Category" },
                { "Extra kategori 1", "Extra Category 1" },
                { "Extra kategori 2", "Extra Category 2" },
                { "Sortering", "Sorting" },
                { "Dold", "Hidden" },
                { "Aktiv", "Active" },
                { "Leverantör", "Supplier" },
                { "ShopId 278 Driva Lehmanns AB", "ShopId 278 Driva Lehmanns AB" },
                { "ShopId 1 fronta.se", "ShopId 1 fronta.se" }
            };

            // Get the English headers names from the second file
            var englishHeaders = GetColumnHeaders(worksheet2);

            // Create a new Excel package for the output
            using var packageOutput = new ExcelPackage();
            var worksheetOutput = packageOutput.Workbook.Worksheets.Add("MatchedRecords");

            // Write the Swedish headers to the output worksheet
            int headerIndex = 1;
            foreach (var header in translationDictionary.Keys)
            {
                worksheetOutput.Cells[1, headerIndex++].Value = header;
            }

            // Write the matched records to the output file
            int outputRow = 2;
            bool duplicateRowInserted = false; // Flag to track the insertion of the duplicate row
            string previousTJStyleNo = ""; // Variable to store the previous value of "TJ_Style_no"

            foreach (int row in Enumerable.Range(2, worksheet2.Dimension.End.Row - 1))
            {
                // Check if the value of "TJ_Style_no" has changed
                string currentTJStyleNo = worksheet2.Cells[row, englishHeaders.IndexOf("TJ_Style_no") + 1].Value?.ToString();

                // Insert the duplicate row once immediately after headers
                if (!duplicateRowInserted)
                {
                    InsertDuplicateRow(worksheetOutput, outputRow, translationDictionary, hardcodedValues, groupToHuvudkategori, worksheet2, englishHeaders, row);
                    outputRow++; // Move to the next row after the duplicate row
                    duplicateRowInserted = true; // Set the flag to true after the first insertion
                }
                else if (currentTJStyleNo != previousTJStyleNo && !string.IsNullOrEmpty(previousTJStyleNo))
                {
                    // Insert a duplicate row before the change with specific columns set to null
                    InsertDuplicateRow(worksheetOutput, outputRow, translationDictionary, hardcodedValues, groupToHuvudkategori, worksheet2, englishHeaders, row);
                    outputRow++; // Move to the next row after the duplicate row
                }

                // Write the data to the current row
                int outputCol = 1;
                foreach (var swedishHeader in translationDictionary.Keys)
                {
                    var englishHeader = translationDictionary.ContainsKey(swedishHeader) ? translationDictionary[swedishHeader] : null;
                    if (englishHeader == "Weight")
                    {
                        worksheetOutput.Cells[outputRow, outputCol].Value = null;
                    }
                    else if (swedishHeader == "Pris")
                    {
                        worksheetOutput.Cells[outputRow, outputCol].Value = null;
                    }
                    else if (swedishHeader == "Sortering")
                    {
                        worksheetOutput.Cells[outputRow, outputCol].Value = "0"; // Set Sortering column to "0"
                    }
                    else if (swedishHeader.StartsWith("Huvudbildadress") || swedishHeader.StartsWith("Kompleterandebildadress"))
                    {
                        worksheetOutput.Cells[outputRow, outputCol].Value = TransformUrl(worksheet2.Cells[row, englishHeaders.IndexOf(englishHeader) + 1].Value?.ToString());
                    }
                    else if (englishHeader != null && englishHeaders.Contains(englishHeader))
                    {
                        int colIndex = englishHeaders.IndexOf(englishHeader) + 1;
                        worksheetOutput.Cells[outputRow, outputCol].Value = worksheet2.Cells[row, colIndex].Value;
                    }
                    outputCol++;
                }

                outputRow++; // Move to the next row
                previousTJStyleNo = currentTJStyleNo; // Update the previous value
            }

            // Save the output file
            packageOutput.SaveAs(new FileInfo(outputPath));

            Console.WriteLine($"Matched columns and data have been written to the output file: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <summary>
    /// Inserts a duplicate row with specific columns set to null or hardcoded values.
    /// </summary>
    /// <param name="worksheet">The target worksheet.</param>
    /// <param name="rowIndex">The index of the row to insert.</param>
    /// <param name="translationDictionary">The translation dictionary for headers.</param>
    /// <param name="hardcodedValues">The hardcoded values dictionary.</param>
    /// <param name="groupToHuvudkategori">The group to Huvudkategori mapping dictionary.</param>
    /// <param name="worksheet2">The source worksheet.</param>
    /// <param name="englishHeaders">The list of English headers.</param>
    /// <param name="sourceRow">The index of the source row.</param>
    static void InsertDuplicateRow(ExcelWorksheet worksheet, int rowIndex, Dictionary<string, string> translationDictionary, Dictionary<string, string> hardcodedValues, Dictionary<string, string> groupToHuvudkategori, ExcelWorksheet worksheet2, List<string> englishHeaders, int sourceRow)
    {
        worksheet.InsertRow(rowIndex, 1); // Insert a new row at the specified index
        bool firstExtraKategori1 = true; // Flag to check the first occurrence of "Extra kategori 1"

        // Check the "Group" value
        string groupValue = worksheet2.Cells[sourceRow, englishHeaders.IndexOf("Group") + 1].Value?.ToString().ToUpper();

        for (int col = 1; col <= translationDictionary.Count; col++)
        {
            string header = worksheet.Cells[1, col].Value?.ToString();
            string englishHeader = translationDictionary.ContainsKey(header) ? translationDictionary[header] : null;

            if (header == "Artikelnr variant" || header == "Storlek" || header == "Färg" || header == "Lager")
            {
                worksheet.Cells[rowIndex, col].Value = null; // Set specific columns to null
            }
            else if (header == "Vikt")
            {
                var weightValue = worksheet2.Cells[sourceRow, englishHeaders.IndexOf("Weight") + 1].Value?.ToString();
                if (!string.IsNullOrEmpty(weightValue))
                {
                    // Extract numeric value from the string (remove 'g' suffix)
                    var numericValue = Regex.Match(weightValue, @"\d+").Value;
                    worksheet.Cells[rowIndex, col].Value = numericValue;
                }
                else
                {
                    worksheet.Cells[rowIndex, col].Value = null;
                }
            }
            else if (header == "Extra kategori 1")
            {
                if (firstExtraKategori1)
                {
                    worksheet.Cells[rowIndex, col].Value = hardcodedValues.ContainsKey(header) ? hardcodedValues[header] : null;
                    firstExtraKategori1 = false; // After setting value for the first occurrence, set the flag to false
                }
                else
                {
                    worksheet.Cells[rowIndex, col].Value = null; // Set null for subsequent occurrences
                }
            }
            else if (header == "Sortering")
            {
                worksheet.Cells[rowIndex, col].Value = "0"; // Set Sortering column to "0"
            }
            else if (header == "Huvudkategori" && !string.IsNullOrEmpty(groupValue) && groupToHuvudkategori.ContainsKey(groupValue))
            {
                worksheet.Cells[rowIndex, col].Value = groupToHuvudkategori[groupValue]; // Set Huvudkategori based on Group value
            }
            else if (hardcodedValues.ContainsKey(header))
            {
                worksheet.Cells[rowIndex, col].Value = hardcodedValues[header]; // Use hardcoded value if applicable
            }
            else if (englishHeader != null && englishHeaders.Contains(englishHeader))
            {
                int colIndex = englishHeaders.IndexOf(englishHeader) + 1;
                if (englishHeader.StartsWith("Model") || englishHeader.StartsWith("Front") || englishHeader.StartsWith("Back") || englishHeader.StartsWith("Left"))
                {
                    // Apply URL transformation for image columns
                    worksheet.Cells[rowIndex, col].Value = TransformUrl(worksheet2.Cells[sourceRow, colIndex].Value?.ToString());
                }
                else
                {
                    worksheet.Cells[rowIndex, col].Value = worksheet2.Cells[sourceRow, colIndex].Value; // Copy other columns from the source row
                }
            }
        }
    }

    /// <summary>
    /// Transforms a URL to the required format.
    /// </summary>
    /// <param name="url">The original URL.</param>
    /// <returns>The transformed URL.</returns>
    static string TransformUrl(string url)
    {
        if (string.IsNullOrEmpty(url))
            return null;

        var match = Regex.Match(url, @"([^/]+/[^/]+\.[^/]+)$");
        if (match.Success)
        {
            // Remove "Img/" from the matched URL
            var cleanedUrl = match.Value.Replace("Img/", "");
            return $"Product/{cleanedUrl}";
        }
        return url;
    }

    /// <summary>
    /// Gets the column headers from a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to extract headers from.</param>
    /// <returns>A list of column headers.</returns>
    static List<string> GetColumnHeaders(ExcelWorksheet worksheet)
    {
        var headers = new List<string>();
        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
        {
            var headerValue = worksheet.Cells[1, col].Value?.ToString();
            if (!string.IsNullOrEmpty(headerValue))
            {
                headers.Add(headerValue);
            }
        }
        return headers;
    }

    /// <summary>
    /// Converts a CSV file to Excel format if necessary.
    /// </summary>
    /// <param name="filePath">The file path of the input file.</param>
    /// <returns>The file path of the converted or original Excel file.</returns>
    static string ConvertToExcelIfCSV(string filePath)
    {
        if (Path.GetExtension(filePath).Equals(".csv", StringComparison.OrdinalIgnoreCase))
        {
            string excelFilePath = Path.ChangeExtension(filePath, ".xlsx");

            // Create Excel package
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Read CSV data using CsvHelper
                using (var reader = new StreamReader(filePath))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    int rowIndex = 1;
                    while (csv.Read())
                    {
                        for (int col = 0; col < csv.Parser.Record.Length; col++)
                        {
                            // Write CSV data to Excel worksheet
                            worksheet.Cells[rowIndex, col + 1].Value = csv.GetField(col);
                        }
                        rowIndex++;
                    }
                }

                // Save Excel package
                package.SaveAs(new FileInfo(excelFilePath));
            }

            // Return the path to the newly created Excel file
            return excelFilePath;
        }

        // Return original file path if not a CSV
        return filePath;
    }

    /// <summary>
    /// Generates a unique file name by appending a number if a file with the same name exists.
    /// </summary>
    /// <param name="directory">The directory where the file will be saved.</param>
    /// <param name="baseFileName">The base file name.</param>
    /// <returns>A unique file name with the directory path.</returns>
    static string GenerateUniqueFileName(string directory, string baseFileName)
    {
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(baseFileName);
        string extension = Path.GetExtension(baseFileName);

        int count = 1;
        string fullPath;
        do
        {
            string tempFileName = count == 1 ? baseFileName : $"{fileNameWithoutExtension}_{count}{extension}";
            fullPath = Path.Combine(directory, tempFileName);
            count++;
        } while (File.Exists(fullPath));

        return fullPath;
    }
}
