using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.IO;

using System.Security.Cryptography.X509Certificates;

using Microsoft.Win32;

using OfficeOpenXml;






namespace alarmVerification
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

#pragma warning disable IDE0059

#pragma warning disable CS8600

#pragma warning disable CS8604





    public partial class MainWindow : Window
    {

        Globals g = new();


        WPFclass wpf = new WPFclass();

        public MainWindow()
        {
            InitializeComponent();

            Globals.dblGonToRad = 0.0157079633;
            Globals.LeftRailCol = 1;
            Globals.RightRailCol = 4;
            Globals.iFirstDataRow = 7;
            Globals.TracksWorksheet = "Tracks";


            // Console settings
            Console.OutputEncoding = System.Text.Encoding.Unicode;

            // Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

        }





        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void tbAlarmTargetName_TextChanged(object sender, TextChangedEventArgs e)
        {
            //btnSelectProjectWorkbook.IsEnabled= true;   
        }

        private void tbWorkbookFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {

        }



        private void ListBoxItem_Selected(object sender, RoutedEventArgs e)
        {

        }

        private void lbDaysOfHistoricData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void btnSelectProjectWorkbook_Click(object sender, RoutedEventArgs e)
        {


            var ofd = new Microsoft.Win32.OpenFileDialog() { Filter = "Excel Files (*.xlsx)|*.xlsx" };
            var result = ofd.ShowDialog();
            if (result == false)
            {
                return;
            }

            string strSelectedFile = ofd.FileName;
            tbWorkbookFilePath.Text = strSelectedFile;

            // activate the buttons
            btnGenerateReport.IsEnabled = true;

        }

        private void btnGenerateReport_Click(object sender, RoutedEventArgs e)
        {
            string strExcelFilePath = tbWorkbookFilePath.Text; 
            string strSelectedProject = tbSelectedProject.Text;
            string strDaysOfHistoricData = tbSelectedDaysOfHistoricData.Text;
            string strPrismsBracketing = tbSelectedPrismsBracketing.Text;
            string strTracksWorksheet = Globals.TracksWorksheet;
            string strFirstDataRow = Globals.iFirstDataRow.ToString();
            string strAlarmPrism = tbAlarmTargetName.Text;

            int iNoOfPrisms = wpf.countPrisms(strExcelFilePath, strTracksWorksheet, strFirstDataRow,1);

            Globals.iLastDataRow = Globals.iFirstDataRow + iNoOfPrisms-1;
            string strLastDataRow = Globals.iLastDataRow.ToString();

            // Locate the Alarm prism
            var RowCol = wpf.locateAlarmPrism(strExcelFilePath, strTracksWorksheet, strFirstDataRow, strLastDataRow, strAlarmPrism);
            int iRow = RowCol.Item1;
            int iCol = RowCol.Item2;

            // Locate the bracketing Rail Start and Rail End rows
            string strAlarmPrismRow = iRow.ToString();
            string strAlarmPrismCol = iCol.ToString();

            var RowStartEnd = wpf.locateRailStartEnd(strExcelFilePath, strTracksWorksheet, strAlarmPrismRow, strAlarmPrismCol);
            int iTrackRowStart = RowStartEnd.Item1;
            int iTrackRowEnd = RowStartEnd.Item2;

            // define the Alarm prism data window

            int iFirstDataBlockRow = iRow - Convert.ToInt16(strPrismsBracketing);
            int iLastDataBlockRow = iRow + Convert.ToInt16(strPrismsBracketing);
            if (iFirstDataBlockRow < iTrackRowStart) { iFirstDataBlockRow = iTrackRowStart; }
            if (iLastDataBlockRow > iTrackRowEnd) { iLastDataBlockRow = iTrackRowEnd; }

            // define the start date and the end date of the data window.

            int iDays = (Convert.ToInt32(strDaysOfHistoricData)-1) * -1;
            string strTimeBlockEnd = " '" + DateTime.Now.ToString("yyyy-MM-dd") + " 23:59:00' ";
            string strTimeBlockStart = " '" + DateTime.Now.AddDays(iDays).ToString("yyyy-MM-dd")+ " 00:00:01' ";



            MessageBox.Show(strTimeBlockStart+"\n"+strTimeBlockEnd);

        }
    }
}
