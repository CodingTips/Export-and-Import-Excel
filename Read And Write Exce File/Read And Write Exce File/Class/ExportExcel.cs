
using Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class WorkWithExcel
{

    public string CreateExcelFile(DirectoryInfo outputDir)
    {

        FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
        if (newFile.Exists)
        {
            newFile.Delete();  // ensures we create a new workbook
            newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
        }
        using (ExcelPackage package = new ExcelPackage(newFile))
        {
            // add a new worksheet to the empty workbook

            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Student Lists");
            //Here  column to  the current sheet
            worksheet.Cells[1, 1].Value = "Title";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "Family";
            worksheet.Cells[1, 4].Value = "Student Code";


            //Add value to celles

            worksheet.Cells[2, 1].Value = "Mr.";
            worksheet.Cells[2, 2].Value = "Mehran";
            worksheet.Cells[2, 3].Value = "Janfeshan";
            worksheet.Cells[2, 4].Value = "ST54516";



            worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells
            worksheet.HeaderFooter.OddFooter.RightAlignedText =
                string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
            // add the sheet name to the footer
            worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
            //Save and export excel file

            package.Save();

        }

        return newFile.FullName;
    }



    public DataSet ReadExcel()
    {
        string filePath;
        filePath = "C:\\sample1.xlsx";
        FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        excelReader.IsFirstRowAsColumnNames = true;
        DataSet result = excelReader.AsDataSet();
        excelReader.Close();
        return result;
    }



}

