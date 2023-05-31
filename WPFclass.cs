using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using OfficeOpenXml;

#pragma warning disable IDE0059

#pragma warning disable CS8600

#pragma warning disable CS8604




namespace alarmVerification
{
    public partial class WPFclass
    {


        public int countPrisms(string strExcelWorkbookFullPath, string strWorksheet, string strFirstDataRow, int iColumnToCheck)
        {

            //
            // Purpose:
            //      To count the number of rail prisms in the project
            // Input:
            //      Full path and name of workbook
            //      Name of reference worksheet
            // Output:
            //      Number of prisms
            // Useage:
            //      int iNoOfPrisms = gnaSpreadsheetAPI.countPrisms(strMasterWorkbookFullPath, strReferenceWorksheet, strFirstDataRow);
            //

            int iPrisms = 0;
            string strName = "x";
            int iRow = Convert.ToInt16(strFirstDataRow);

            FileInfo excelWorkbook = new(strExcelWorkbookFullPath);

            using (ExcelPackage package = new(excelWorkbook))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];
                do
                {
                    ++iPrisms;
                    ++iRow;
                    strName = Convert.ToString(namedWorksheet.Cells[iRow, iColumnToCheck].Value);
                } while (strName != "");

                try
                {
                    package.Dispose();
                }
                catch (Exception ex)
                {
                    errorMessage(ex);
                }

                return iPrisms;

            }
        }

        public void errorMessage(Exception ex)
        {
            string strErrorMessage = "Error:\nWorkbook open or missing\n" + ex;
            MessageBox.Show(strErrorMessage);
            Environment.Exit(0);
        }


        public Tuple<int, int> locateAlarmPrism(string strExcelWorkbookFullPath, string strWorksheet, string strFirstDataRow, string strLastDataRow, string strAlarmPrism)
        {

            int iFirstRow = Convert.ToInt16(strFirstDataRow);
            int iLastRow = Convert.ToInt16(strLastDataRow);

            int iCol = 0, iRow = 0;




            FileInfo excelWorkbook = new(strExcelWorkbookFullPath);

            using (ExcelPackage package = new(excelWorkbook))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strWorksheet];
                iCol = Globals.LeftRailCol;
                for (int i = 1; i <= 2; i++)
                {
                    for (iRow = iFirstRow; iRow < iLastRow; iRow++)
                    {
                        string strName = Convert.ToString(namedWorksheet.Cells[iRow, iCol].Value);
                        if (strName == strAlarmPrism) { goto ExitPoint; }
                    }
                    iCol = Globals.RightRailCol;
                }

                iRow = -99;
                iCol = -99;

                try
                {
                    package.Dispose();
                }
                catch (Exception ex)
                {
                    errorMessage(ex);
                }

            }

ExitPoint:
            return new Tuple<int, int>(iRow, iCol);
        }



        public Tuple<int, int> locateRailStartEnd(string strExcelFilePath, string strTracksWorksheet, string strAlarmPrismRow, string strAlarmPrismCol)
        {
            int iRow = Convert.ToInt16(strAlarmPrismRow);
            int iCol = Convert.ToInt16(strAlarmPrismCol);
            int iAlarmPrismRow = iRow;
            int iTrackRowStart = 0;
            int iTrackRowEnd = 0;

            FileInfo excelWorkbook = new(strExcelFilePath);

            using (ExcelPackage package = new(excelWorkbook))
            {
                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strTracksWorksheet];
                iCol = Globals.LeftRailCol + 1;
                // backwards first to find start
                for (iRow = iAlarmPrismRow; iRow > 0; iRow--)
                {
                    string strTrackStart = Convert.ToString(namedWorksheet.Cells[iRow, iCol].Value);
                    if (strTrackStart == "Rail Start") { goto ExitPoint1; }
                    if (strTrackStart == "Rail start") { goto ExitPoint1; }
                    if (strTrackStart == "Rail_start") { goto ExitPoint1; }
                    if (strTrackStart == "Rail_Start") { goto ExitPoint1; }
                }

ExitPoint1:
                iTrackRowStart = iRow;

                // Forwards first to find end
                for (iRow = iAlarmPrismRow; iRow < 3000; iRow++)
                {
                    string strTrackStart = Convert.ToString(namedWorksheet.Cells[iRow, iCol].Value);
                    if (strTrackStart == "Rail End") { goto ExitPoint2; }
                    if (strTrackStart == "Rail end") { goto ExitPoint2; }
                    if (strTrackStart == "Rail_end") { goto ExitPoint2; }
                    if (strTrackStart == "Rail_End") { goto ExitPoint2; }
                }

ExitPoint2:
                iTrackRowEnd = iRow;


                try
                {
                    package.Dispose();
                }
                catch (Exception ex)
                {
                    errorMessage(ex);
                }

            }

            return new Tuple<int, int>(iTrackRowStart, iTrackRowEnd);

        }






    }
}
