using System;
using System.IO;
using System.Linq;
using Tesseract;
using ClosedXML.Excel;
using System.Drawing;

namespace Medical_Text_Extracter
{
    public class Extractor
    {

        private static TesseractEngine tesseractEngine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default);

        // Preconditions: image_path is the Absolute File Path to the image.
        // Postconditions: Returns a string array with all text on the image.
        public static string GetText(string image_path)
        {
            try
            {
                using (Pix img = Pix.LoadFromFile(image_path))
                {
                    using (Page page = tesseractEngine.Process(img))
                    {
                        return page.GetText();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        //
        public static string GetEdgeText(string image_path, string temp_save_path)
        {

            string top_text = null;
            string bottom_text = null;

            try
            {
                using (Bitmap origImage = new Bitmap(image_path))
                {
                    Rectangle cropAreaTop = new Rectangle((int)(origImage.Width * 0.1), 0, (int)(origImage.Width * 0.8), (int)(origImage.Height * 0.1));
                    Rectangle cropAreaBottom = new Rectangle((int)(origImage.Width * 0.1), (int)(origImage.Height * 0.9), (int)(origImage.Width * 0.8), (int)(origImage.Height * 0.1));

                    using (Bitmap croppedImageTop = origImage.Clone(cropAreaTop, origImage.PixelFormat))
                    {
                        croppedImageTop.Save(temp_save_path);
                        top_text = GetText(temp_save_path);
                        File.Delete(temp_save_path);
                    }
                    using (Bitmap croppedImageBottom = origImage.Clone(cropAreaBottom, origImage.PixelFormat))
                    {
                        croppedImageBottom.Save(temp_save_path);
                        bottom_text = GetText(temp_save_path);
                        File.Delete(temp_save_path);
                    }
                }
            }
            catch (Exception e)
            {
                return null;
            }

            if (string.IsNullOrEmpty(top_text) && string.IsNullOrEmpty(bottom_text))
            {
                string allText = GetText(image_path);
                return allText;
            }
            else if (string.IsNullOrEmpty(top_text))
            {
                return bottom_text;
            }
            else if (string.IsNullOrEmpty (bottom_text))
            {
                return top_text;
            }
            else
            {
                return null;
            }
        }



        // Preconditions: Excel_File_Path is the correct file path to the .xlsm file, image_path is the absolute path to a singular image, classification is it's correct classification, row is greater than one, worksheet number is greater than or equal to one.
        // Postconditions: This will set the cells in specified row with the image path and it's classification on the corresponding worksheet of its .xlsm file.
        public static void PutImageIntoExcelDoc(string Excel_File_path, string image_path, string temp_save_path, int row = 1, int worksheet_number = 1) 
        {
            using (XLWorkbook workbook = new XLWorkbook(Excel_File_path)) 
            {
                if (worksheet_number < 1 || worksheet_number > workbook.Worksheets.Count) throw new ArgumentException("InvalidWorksheetNumberError");
                var worksheet = workbook.Worksheet(worksheet_number);
                if (row <= 1) row = worksheet.LastRowUsed().RowNumber() + 1;
                string classification = GetEdgeText(image_path, temp_save_path);
                worksheet.Cell(row, 1).Value = image_path;
                worksheet.Cell(row, 2).Value = classification;
                workbook.Save();
            }
        }


        // Preconditions: file_path is the absolute file path in which all of the .tiff files are stored, ExcelFileSavePath is the path the user would like to use to save the newly created Excel file there.
        // Postconditions: A new Excel file will be created, being saved at the specified path. It will have two columns, one being the absolute file path to an image, and the second being a string representing its classification.
        public static void PutAllImagesIntoExcelDoc(string file_path, string ExcelFileSavePath, string temp_save_path)
        {
            using(XLWorkbook workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Image Classification");
                worksheet.Cell(1, 1).Value = "Image File Path";
                worksheet.Cell(1, 2).Value = "Image Classification";
                workbook.SaveAs(ExcelFileSavePath);

                int row = 2;
                string[] tiffFiles = Directory.GetFiles(file_path);
                foreach (var tiffFile in tiffFiles)
                {
                    string classification = GetEdgeText(tiffFile, temp_save_path);
                    worksheet.Cell(row, 1).Value = tiffFile;
                    worksheet.Cell(row, 2).Value = classification;
                    ++row;
                }
                workbook.Save();
            }
        }


        
   
        private static string ProcessText(string text)
        {
            return "";
        }

        public static bool AlphaSort(string text_doc_file_path)
        {
            try
            {
                string[] terms = File.ReadAllLines(text_doc_file_path);
                Array.Sort(terms, StringComparer.InvariantCultureIgnoreCase);
                File.WriteAllLines(text_doc_file_path, terms);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    
        
    }
}
