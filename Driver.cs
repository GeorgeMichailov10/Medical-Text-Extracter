using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Medical_Text_Extracter
{
    class Driver
    {
        static void Main(string[] args)
        {
            string images_path = @"C:\path\to\tiff\files";
            string temp_save_path = @"C:\path\to\temp\file\tmp.tif";
            string excel_file_save_path = @"C:\path\to\save\excel\file\Classifications.xlsx";

            try
            {
                // Process all TIFF images and save the results in an Excel file
                Extractor.PutAllImagesIntoExcelDoc(images_path, excel_file_save_path, temp_save_path);
                Console.WriteLine("All images have been processed and saved in the Excel file.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
