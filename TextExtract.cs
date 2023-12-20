using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Tesseract;
using ClosedXML.Excel;
using System.Drawing;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Drawing.Imaging;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.Generic;

namespace Medical_Text_Extracter
{
    public class Extractor
    {

        private static readonly object ExcelLock = new object();

        // Expects Path to images, path and name to where you want to save the Excel file, and a (prefereably empty) folder where it will store temporary crops 
        public static void ProcessAllFiles(string files_path, string excel_file_save_path, string temp_cropping_save_file)
        {
            CreateExcelDocument(excel_file_save_path);
            ProcessAllTiffFilesSequentially(files_path, excel_file_save_path, temp_cropping_save_file);
        }


        // Creates the Excel Document where everything will be saved.
        private static void CreateExcelDocument(string excel_file_save_path)
        {
            using (XLWorkbook workbook = new XLWorkbook())
            {
                // Create sheets for each of the different types of files.
                workbook.Worksheets.Add("Tiff File Classification");
                workbook.Worksheets.Add("MP4 File Classification");
                workbook.SaveAs(excel_file_save_path);
            }
        }


        // FUTURE IMPLEMENTATION FOR EFFICIENCY
        private static async Task ProcessAllTiffFiles(string files_path, string excel_file_save_path, string temp_cropping_save_file)
        {
            using (XLWorkbook workbook = new XLWorkbook(excel_file_save_path))
            {
                var worksheet = workbook.Worksheet(1);
                worksheet.Cell(1, 1).Value = "Tiff File Path";
                worksheet.Cell(1, 2).Value = "Image Classification";

                int row = 2;
                string[] tiffFiles = Directory.GetFiles(files_path, "*.tiff");
                var tasks = new List<Task>();

                foreach (string tiffFile in tiffFiles)
                {
                    int curr_row = row;
                    Task task = Task.Run(() => ProcessTiffFile(tiffFile, worksheet, curr_row, temp_cropping_save_file));
                    tasks.Add(task);
                    ++row;
                }

                await Task.WhenAll(tasks);
                workbook.Save();
            }
        }

        // FUTURE IMPLEMENTATION FOR EFFICIENCY
        private static async Task ProcessTiffFile(string image_path, IXLWorksheet worksheet, int row, string temp_cropping_save_path)
        {
            string classification;
            string bottom_text = GetBottomText(image_path, temp_cropping_save_path);
            if (bottom_text != "Failure")
            {
                classification = GetClassification(bottom_text);
                if (classification != null)
                {
                    lock (ExcelLock)
                    {
                        worksheet.Cell(row, 1).Value = image_path;
                        worksheet.Cell(row, 2).Value = classification;
                    }
                    return;
                }
            }
            string left_center_text = GetLeftCenterText(image_path, temp_cropping_save_path);
            if (left_center_text != "Failure")
            {
                classification = GetClassification(left_center_text);
                if (classification != null)
                {
                    lock (ExcelLock)
                    {
                        worksheet.Cell(row, 1).Value = image_path;
                        worksheet.Cell(row, 2).Value = classification;
                    }
                    return;
                }
            }
            string all_text = GetAllText(image_path);
            if (all_text != "Failure")
            {
                classification = GetClassification(all_text);
                if (classification != null)
                {
                    lock (ExcelLock)
                    {
                        worksheet.Cell(row, 1).Value = image_path;
                        worksheet.Cell(row, 2).Value = classification;
                    }
                    return;
                }
            }
            else
            {
                lock (ExcelLock)
                {
                    worksheet.Cell(row, 1).Value = image_path;
                    worksheet.Cell(row, 2).Value = "Failure";
                }
            }
        }


        // Handles Processing all tiff files
        private static void ProcessAllTiffFilesSequentially(string files_path, string excel_file_save_path, string temp_cropping_save_file)
        {
            using (XLWorkbook workbook = new XLWorkbook(excel_file_save_path))
            {
                var worksheet = workbook.Worksheet(1);
                worksheet.Cell(1, 1).Value = "Tiff File Path";
                worksheet.Cell(1, 2).Value = "Image Classification";

                int row = 2;
                string[] tiffFiles = Directory.GetFiles(files_path, "*.tiff");

                foreach (string tiffFile in tiffFiles)
                {
                    ProcessTiffFileSequentially(tiffFile, worksheet, row, temp_cropping_save_file);
                    ++row;
                }

                workbook.Save();
            }
        }

        // Processes one tiff file as part of the sequential process.
        private static async Task ProcessTiffFileSequentially(string image_path, IXLWorksheet worksheet, int row, string temp_cropping_save_path)
        {
            string classification;
            string bottom_text = GetBottomText(image_path, temp_cropping_save_path);
            if (bottom_text != "Failure")
            {
                classification = GetClassification(bottom_text);
                if (classification != null)
                {
                    lock (ExcelLock)
                    {
                        worksheet.Cell(row, 1).Value = image_path;
                        worksheet.Cell(row, 2).Value = classification;
                    }
                    return;
                }
            }
            string left_center_text = GetLeftCenterText(image_path, temp_cropping_save_path);
            if (left_center_text != "Failure")
            {
                classification = GetClassification(left_center_text);
                if (classification != null)
                {
                    lock (ExcelLock)
                    {
                        worksheet.Cell(row, 1).Value = image_path;
                        worksheet.Cell(row, 2).Value = classification;
                    }
                    return;
                }
            }
            string all_text = GetAllText(image_path);
            if (all_text != "Failure")
            {
                classification = GetClassification(all_text);
                if (classification != null)
                {
                    lock (ExcelLock)
                    {
                        worksheet.Cell(row, 1).Value = image_path;
                        worksheet.Cell(row, 2).Value = classification;
                    }
                    return;
                }
            }
            else
            {
                lock (ExcelLock)
                {
                    worksheet.Cell(row, 1).Value = image_path;
                    worksheet.Cell(row, 2).Value = "Failure";
                }
            }
        }


        // Creates a crop of bottom portion, gets text, deletes crop, returns raw text.
        private static string GetBottomText(string image_path, string temp_cropping_save_path) 
        {
            string bottom_image_path = CreateImageCropping(image_path, temp_cropping_save_path, 0.2, 0.9, 0.6, 0.1);
            string bottom_text = GetText(bottom_image_path);
            if (bottom_text != "Failure") File.Delete(bottom_image_path);
            return bottom_text.Length == 0 ? "Failure" : bottom_text;
        }

        // Creates a crop of Left Center portion, gets text, deletes crop, returns raw text.
        private static string GetLeftCenterText(string image_path, string temp_cropping_save_path)
        {
            string left_center_image_path = CreateImageCropping(image_path, temp_cropping_save_path, 0.09, 0.17, 0.22, 0.09, true);
            string text = GetText(left_center_image_path);
            if (text != "Failure") File.Delete(left_center_image_path);
            return text.Length == 0 ? "Failure" : text;
        }

        // Returns all text on the image
        private static string GetAllText(string image_path)
        {
            string text = GetText(image_path);
            return text.Length == 0 ? "Failure" : text;
        }

        // Function that creates a cropping based on the given parameters, image path, and returns its file path
        private static string CreateImageCropping(string image_path, string temp_cropping_save_path, double x_percent, double y_percent, double width_percent, double height_percent, bool is_special_case = false)
        {
            try
            {
                using (Bitmap originalImage = new Bitmap(image_path))
                {
                    Rectangle cropArea = new Rectangle((int)(x_percent * originalImage.Width), (int)(y_percent * originalImage.Height), (int)(width_percent * originalImage.Width), (int)(height_percent * originalImage.Height));
                    using (Bitmap croppedImage = new Bitmap(cropArea.Width, cropArea.Height))
                    {
                        using (Graphics g = Graphics.FromImage(croppedImage))
                        {
                            // Draw the specified section of the source image to the new one
                            g.DrawImage(originalImage, 0, 0, cropArea, GraphicsUnit.Pixel);
                        }
                        if (is_special_case)
                        {
                            for (int Height = 0; Height < croppedImage.Height; Height++)
                            {
                                for (int Width = 0; Width < croppedImage.Width; Width++)
                                {
                                    // Get the pixel color
                                    System.Drawing.Color pixelColor = croppedImage.GetPixel(Width, Height);

                                    // Check if the color is neither black nor white
                                    if (pixelColor.R == 128 && pixelColor.G == 255 && pixelColor.B == 128)
                                    {
                                        // Change the color to black
                                        croppedImage.SetPixel(Width, Height, System.Drawing.Color.Black);
                                    }
                                }
                            }
                        }
                        string newFilePath = Path.Combine(Path.GetDirectoryName(image_path), $"{Path.GetFileNameWithoutExtension(image_path)}_temp{Path.GetExtension(image_path)}");
                        croppedImage.Save(newFilePath, System.Drawing.Imaging.ImageFormat.Tiff);
                        return newFilePath;
                    }
                }
            }
            catch (Exception e)
            {
                return "Failure";
            }
        }
        
        // Returns most likely string candidate for tiff Files (Will not work for Videos because there is too much writing)
        private static string GetClassification(string raw_text)
        {
            string[] raw_lines = raw_text.Split(new char[] {'\n'}, StringSplitOptions.RemoveEmptyEntries);
            string[] processed_lines = PreProcessText(raw_lines);
            string longest = null;
            for (int i = processed_lines.Length - 1; i >= 0; i--)
            {
                if (longest == null || longest.Length < processed_lines[i].Length) longest = processed_lines[i];
            }
            return longest.ToLower();
        }

        // Removes any strings that contain anything other than letters and spaces
        private static string[] PreProcessText(string[] raw_lines)
        {
            List<string> processed_lines = new List<string>();
            foreach (string line in raw_lines)
            {
                if (line.All(c => char.IsLetter(c) || char.IsWhiteSpace(c)) && (line.Length >= 2 && line.Substring(0, 2) != "P "))
                {
                    processed_lines.Add(line);
                } 
            }
            return processed_lines.ToArray();
        }

        // Returns all text from an image as a string split into lines.
        private static string GetText(string imagePath)
        {
            try
            {
                using (TesseractEngine tesseractEngine = new TesseractEngine(@"D:\Applications\Tesseract OCR\tessdata", "eng", EngineMode.Default))
                {
                    using (Pix img = Pix.LoadFromFile(imagePath))
                    {
                        using (Tesseract.Page page = tesseractEngine.Process(img))
                        {
                            return page.GetText();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return "Failure";
            }
        }

    }
}

