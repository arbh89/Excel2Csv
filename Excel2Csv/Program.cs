using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace Excel2Csv
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("The syntax of the command is incorrect, type --help for more information");
            }
            else
            {
                if (HelpRequired(args[0]))
                {
                    ShowHelp();
                }
                else if (args.Length >= 2)
                {
                    if (IsValidPath(args[0]))
                    {
                        var folderPath = args[0];
                        var outputPath = args[1];
                        int fileCount = 1;
                        var filesToConvet = new List<string>();
                        var addExtraData = false;
                        //var fileFilterPattern = "*.xlsx";

                        if (File.Exists(folderPath))
                        {
                            filesToConvet.Add(folderPath);
                        }
                        else
                        {
                            filesToConvet.AddRange(Directory.EnumerateFiles(folderPath, "*.xlsx"));
                            filesToConvet.AddRange(Directory.EnumerateFiles(folderPath, "*.xlsm"));
                        }

                        //check if extra data will be added
                        if (args.Length >= 3)
                        {
                            if (args[2] == "/E")
                            {
                                addExtraData = true;
                            }
                            //else if (args[2] == "/P")
                            //{
                            //    fileFilterPattern = args[2];
                            //}
                            //else if (args.Length == 4)
                            //{
                            //    if (args[3] == "/E")
                            //    {
                            //        addExtraData = true;
                            //    }
                            //}
                            else
                            {
                                Console.WriteLine("The syntax of the command is incorrect, type --help for more information");
                                return;
                            }
                        }

                        //If output directory does not exists it gets created
                        if (!Directory.Exists(outputPath))
                        {
                            FileInfo file = new System.IO.FileInfo(outputPath);
                            file.Directory.Create();
                        }

                        foreach (string file in filesToConvet)
                        {
                            var excelFile = new FileInfo(file);
                            using (var package = new ExcelPackage(excelFile))
                            {
                                foreach (var item in package.Workbook.Worksheets)
                                {
                                    if (!item.IsWorksheetEmpty())
                                    {
                                        byte[] csvFile;
                                        if (addExtraData)
                                        {
                                            csvFile = item.ConvertToCsv(file);
                                        }
                                        else
                                        {
                                            csvFile = item.ConvertToCsv();
                                        }

                                        File.WriteAllBytes(outputPath + "\\" + fileCount + " - " + item.Name + ".csv", csvFile);
                                        fileCount++;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("The syntax of the command is incorrect, type --help for more information");
                }
            }
        }

        private static bool IsValidPath(string path)
        {
            var isValid = true;
            try
            {
                var fullPath = System.IO.Path.GetFullPath(path);
            }
            catch (Exception)
            {
                isValid = false;
            }
            return isValid && (System.IO.Directory.Exists(path) || System.IO.File.Exists(path));
        }

        private static bool HelpRequired(string param)
        {
            return param == "-h" || param == "--help" || param == "/?";
        }

        private static void ShowHelp()
        {
            Console.WriteLine(@"
Converts every worksheet from excel files (.xlsx) to csv

excel2csv source destination

source       Specifies the file or folder to be converted.
destination  Specifies the directory for the new file(s). If not exists them is created
filter       Indicates the patter to select files in folder (*.xlsx)
/E           Indicates if extra data like file name and sheet name will be added to the output file");
        }
    }
}