using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using Tesseract;

namespace Medical_Text_Extracter
{
    class Driver
    {
        static void Main(string[] args)
        {
            string images_path = @"D:\iMorgon\Testing\image samples";
            string temp_save_path = @"D:\iMorgon\Testing\image samples\temp";
            string excel_file_save_path = @"D:\iMorgon\Testing\image samples\answers.xlsx";

            try
            {
                // Process all TIFF images and save the results in an Excel file
                Extractor.ProcessAllFiles(images_path, excel_file_save_path, temp_save_path);
                Console.WriteLine("All images have been processed and saved in the Excel file.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
