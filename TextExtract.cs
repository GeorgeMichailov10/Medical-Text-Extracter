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
        // Preconditions: image_path is the Absolute File Path to the image.
        // Postconditions: Returns a string array with all text on the image.
        public static string GetText(string image_path)
        {
            try
            {
                using (TesseractEngine tess = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
                {
                    using (Pix img = Pix.LoadFromFile(image_path))
                    {
                        using (Page page = tess.Process(img))
                        {
                            return page.GetText();
                        }
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
        public static string GetEdgeText(string image_path, string top_image_path, string bottom_image_path)
        {
            if (!CreateCroppedImages(image_path, top_image_path, bottom_image_path))
            {
                CleanUpCropping(top_image_path, bottom_image_path);
                return null;
            }

            string top_text = GetText(top_image_path);
            string bottom_text = GetText(bottom_image_path);
            CleanUpCropping(top_image_path, bottom_image_path);

            if (string.IsNullOrEmpty(top_text) && string.IsNullOrEmpty(bottom_text))
            {
                string allText = GetText(image_path);
                //// TODO: PROCESS TEXT IF YOU CAN
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
                
            }
        }



        // Preconditions: Excel_File_Path is the correct file path to the .xlsm file, image_path is the absolute path to a singular image, classification is it's correct classification, row is greater than one, worksheet number is greater than or equal to one.
        // Postconditions: This will set the cells in specified row with the image path and it's classification on the corresponding worksheet of its .xlsm file.
        public static void PutImageIntoExcelDoc(string Excel_File_path, string image_path, string classification, int row, int worksheet_number = 1) 
        {
            using (XLWorkbook workbook = new XLWorkbook("File path to workbook")) 
            {
                if (worksheet_number < 1 || worksheet_number > workbook.Worksheets.Count) throw new ArgumentException("InvalidWorksheetNumberError");
                var worksheet = workbook.Worksheet(worksheet_number);
                worksheet.Cell(row, 1).Value = image_path;
                worksheet.Cell(row, 2).Value = classification;
                workbook.Save();
            }
        }


        // Preconditions: file_path is the absolute file path in which all of the .tiff files are stored, ExcelFileSavePath is the path the user would like to use to save the newly created Excel file there.
        // Postconditions: A new Excel file will be created, being saved at the specified path. It will have two columns, one being the absolute file path to an image, and the second being a string representing its classification.
        public static void PutAllImagesIntoExcelDoc(string file_path, string ExcelFileSavePath)
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
                    string classification = RemoveNoiseWords(GetText(tiffFile));
                    worksheet.Cell(row, 1).Value = tiffFile;
                    worksheet.Cell(row, 2).Value = classification;
                    workbook.Save();
                    ++row;
                }
                
            }
        }

        // Preconditions: image_path is the absolute file path to an image, top_image_path and bottom_image_path are the locations where client wants croppings to be temporarily saved.
        private static bool CreateCroppedImages(string image_path, string top_image_path, string bottom_image_path)
        {
            try
            {
                Bitmap origImage = new Bitmap(image_path);
                Rectangle cropAreaTop = new Rectangle((int)(origImage.Width * 0.1), 0, (int)(origImage.Width * 0.8), (int)(origImage.Height * 0.1));
                Rectangle cropAreaBottom = new Rectangle((int)(origImage.Width * 0.1), (int)(origImage.Height * 0.9), (int)(origImage.Width * 0.8), (int)(origImage.Height * 0.1));

                Bitmap croppedImageTop = origImage.Clone(cropAreaTop, origImage.PixelFormat);
                Bitmap croppedImageBottom = origImage.Clone(cropAreaBottom, origImage.PixelFormat);

                croppedImageTop.Save(image_path);
                croppedImageBottom.Save(image_path);
                return true;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }
        
        // Preconditions: None.
        // Postconditions: If files exist at those paths, it will delete them.
        private static void CleanUpCropping(string top_image_path, string bottom_image_path)
        {
            try {       File.Delete(top_image_path);        }
            catch (Exception e) {}
            try {       File.Delete(bottom_image_path);        }
            catch (Exception e) { }
        }
   
        private static string ProcessText(string text)
        {

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
