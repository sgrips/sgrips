using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Bartolini.EasySped.Barcode;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace ExcelAddInBarcode
{
    public partial class BarcodeForExcel
    {
        private void RibbonBarcode_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnBarcode_Click(object sender, RibbonControlEventArgs e)
        {

            Window window = e.Control.Context;

            Worksheet activeWorksheet = ((Worksheet)window.Application.ActiveSheet);

            Range  barcodeStrRange = GetPicturePosition(window.Application);



                if (string.IsNullOrWhiteSpace(barcodeStrRange.Cells.Value.ToString ())) return;

                byte[] barcode = BarcodeGenerator.BarcodeImagePNG(barcodeStrRange.Cells.Value.ToString() , 150, 300);


                using (Image image = Image.FromStream(new MemoryStream(barcode)))
                {

                    image.Save("c:\\temp\\output.png", ImageFormat.Png);  // Or Png
                }


                object missing = System.Reflection.Missing.Value;
                Range picPosition = GetPicturePosition(window.Application); // retrieve the range for picture insert
                Pictures p = activeWorksheet.Pictures(missing) as Pictures;
                Picture pic = null;
                pic = p.Insert("c:\\temp\\output.png", missing);
                // pic.Left = Convert.ToDouble(picRange.Left);
                // pic.Top = picRange.Top;
                pic.Placement = XlPlacement.xlFreeFloating; // Can be any of Excel.XlPlacement.XYZ value




        }


        private Range GetPicturePosition(Application  windowsApp)
        {
            return windowsApp.InputBox("Select a Range", "title", Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing, Type.Missing, 8);
        }
    }


}
