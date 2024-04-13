/* Title:           Import Vehicle Info
 * Date:            8-30-17
 * Author:          Terry Holmes */

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
using NewEventLogDLL;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using NewVehicleDLL;
using VehicleInfoDLL;
using DataValidationDLL;

namespace ImportVehicleInfo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        VehicleClass TheVehicleClass = new VehicleClass();
        VehicleInfoClass TheVehicleInfoClass = new VehicleInfoClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();

        //setting up the data
        FindActiveVehicleByBJCNumberDataSet TheFindActiveVehicleByBJCNumberDataSet = new FindActiveVehicleByBJCNumberDataSet();
        FindDOTStatusByStatusDataSet TheFindDOTStatusByStatusDataSet = new FindDOTStatusByStatusDataSet();
        FindGPSStatusByStatusDataSet TheFindGPSStatusByStatusDataSet = new FindGPSStatusByStatusDataSet();
        FindVehicleInfoByBJCNumberDataSet TheFindVehicleInfoByBJCNumberDataSet = new FindVehicleInfoByBJCNumberDataSet();

        ImportedVehicleInfoDataSet TheImportedVehicleInfoDataSet = new ImportedVehicleInfoDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //setting local variables
            string strInformation;
            int intRowCount;
            int intColumnCount;
            int intRowRange;
            int intColumnRange = 0;
            int intVehicleID = 0;
            int intGPSStatusID = 0;
            int intDOTStatus = 0;
            int intBJCNumber;
            int intOdometer;
            string strIMEI;
            string strGPSStatus;
            string strDOTStatus;
            bool blnCDLRequired;
            bool blnMedicalCardRequired;
            int intTamperTag;
            Excel.Application xlWorkOrder;
            Excel.Workbook xlOrderBook;
            Excel.Worksheet xlOrderSheet;
            Excel.Range range;            

            PleaseWait PleaseWait = new PleaseWait();
            PleaseWait.Show();

            try
            {
                xlWorkOrder = new Excel.Application();
                xlOrderBook = xlWorkOrder.Workbooks.Open(@"c:\users\tholmes\desktop\vehicleinfosheet.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlOrderSheet = (Excel.Worksheet)xlWorkOrder.Worksheets.get_Item(1);

                range = xlOrderSheet.UsedRange;
                intRowRange = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                
                
                for (intRowCount = 1; intRowCount <= intRowRange; intRowCount++)
                {
                    ImportedVehicleInfoDataSet.vehicleinfoRow NewVehicleRow = TheImportedVehicleInfoDataSet.vehicleinfo.NewvehicleinfoRow();

                    NewVehicleRow.TransactionDate = DateTime.Now;

                    for (intColumnCount = 1; intColumnCount <= intColumnRange; intColumnCount++)
                    {
                        strInformation = Convert.ToString((range.Cells[intRowCount, intColumnCount] as Excel.Range).Value2);

                        if(intColumnCount == 1)
                        {
                            intBJCNumber = ConvertBJCNumber(strInformation);

                            TheFindActiveVehicleByBJCNumberDataSet = TheVehicleClass.FindActiveVehicleByBJCNumber(intBJCNumber);

                            intVehicleID = TheFindActiveVehicleByBJCNumberDataSet.FindActiveVehicleByBJCNumber[0].VehicleID;

                            NewVehicleRow.BJCNumber = intBJCNumber;
                            NewVehicleRow.VehicleID = intVehicleID;
                        }
                        if(intColumnCount == 2)
                        {
                            strGPSStatus = strInformation;

                            TheFindGPSStatusByStatusDataSet = TheVehicleInfoClass.FindGPSStatusByStatus(strGPSStatus);

                            intGPSStatusID = TheFindGPSStatusByStatusDataSet.FindGPSStatusByStatus[0].GPSStatusID;

                            NewVehicleRow.GPSStatusID = intGPSStatusID;
                        }
                        if (intColumnCount == 3)
                        {
                            strDOTStatus = strInformation;

                            TheFindDOTStatusByStatusDataSet = TheVehicleInfoClass.FindDOTStatusByStatus(strDOTStatus);

                            intDOTStatus = TheFindDOTStatusByStatusDataSet.FindDOTStatusByStatus[0].DOTStatusID;

                            NewVehicleRow.DOTStatusID = intDOTStatus;
                        }
                        if(intColumnCount == 4)
                        {
                            strIMEI = strInformation;

                            NewVehicleRow.IMEI = strIMEI;
                        }
                        if(intColumnCount == 5)
                        {
                            intTamperTag = Convert.ToInt32(strInformation);

                            NewVehicleRow.TamperTag = intTamperTag;
                        }
                        if(intColumnCount == 6)
                        {
                            intOdometer = Convert.ToInt32(strInformation);

                            NewVehicleRow.Odometer = intOdometer;
                        }
                        if(intColumnCount == 7)
                        {
                            if (strInformation == "True")
                                blnMedicalCardRequired = true;
                            else
                                blnMedicalCardRequired = false;

                            NewVehicleRow.MedicalCardRequired = blnMedicalCardRequired;
                        }
                        if (intColumnCount == 8)
                        {
                            if (strInformation == "True")
                                blnCDLRequired = true;
                            else
                                blnCDLRequired = false;

                            NewVehicleRow.CDLRequired = blnCDLRequired;
                        }
                    }

                    TheImportedVehicleInfoDataSet.vehicleinfo.Rows.Add(NewVehicleRow);
                }
                

                xlOrderBook.Close(true, null, null);
                xlWorkOrder.Quit();

                Marshal.ReleaseComObject(xlOrderSheet);
                Marshal.ReleaseComObject(xlOrderBook);
                Marshal.ReleaseComObject(xlWorkOrder);

                dgrResults.ItemsSource = TheImportedVehicleInfoDataSet.vehicleinfo;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Vehicle Info // Main Window // Window Loaded " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

            PleaseWait.Close();
        }
        private int ConvertBJCNumber(string strBJCNumber)
        {
            int intBJCNUmber = 0;
            int intLength;
            char[] chaBJCNumber;
            bool blnIsNotInteger = false;
            string strNewNumber = "";

            try
            {
                //beginning data validation
                blnIsNotInteger = TheDataValidationClass.VerifyIntegerData(strBJCNumber);

                if(blnIsNotInteger == true)
                {
                    chaBJCNumber = strBJCNumber.ToCharArray();

                    strNewNumber = Convert.ToString(chaBJCNumber[4]);
                    strNewNumber += Convert.ToString(chaBJCNumber[5]);
                    strNewNumber += Convert.ToString(chaBJCNumber[6]);
                    strNewNumber += Convert.ToString(chaBJCNumber[7]);

                    intBJCNUmber = Convert.ToInt32(strNewNumber);

                }
                else
                {
                    intBJCNUmber = Convert.ToInt32(strBJCNumber);
                }

                
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Vehicle Info // Main Window // Convert BJC Number " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

            return intBJCNUmber;
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            int intRecordsReturned;
            int intBJCNumber;
            int intVehicleID;
            bool blnCDLRequired;
            bool blnMedicalCardRequired;
            int intDOTStatusID;
            int intGPSStatusID;
            string strIMEI;
            int intTamperTag;
            int intOdometer;

            PleaseWait PleaseWait = new PleaseWait();
            PleaseWait.Show();

            try
            {
                intNumberOfRecords = TheImportedVehicleInfoDataSet.vehicleinfo.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intBJCNumber = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].BJCNumber;
                    intVehicleID = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].VehicleID;
                    intOdometer = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].Odometer;
                    blnCDLRequired = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].CDLRequired;
                    blnMedicalCardRequired = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].MedicalCardRequired;
                    intDOTStatusID = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].DOTStatusID;
                    intGPSStatusID = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].GPSStatusID;
                    intTamperTag = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].TamperTag;
                    strIMEI = TheImportedVehicleInfoDataSet.vehicleinfo[intCounter].IMEI;

                    TheFindVehicleInfoByBJCNumberDataSet = TheVehicleInfoClass.FindVehicleInfoByBJCNumber(intBJCNumber);

                    intRecordsReturned = TheFindVehicleInfoByBJCNumberDataSet.FindVehicleInfoByBJCNumber.Rows.Count;

                    if(intRecordsReturned == 0)
                    {
                        TheVehicleInfoClass.InsertVehicleInfo(intVehicleID, blnCDLRequired, blnMedicalCardRequired, intDOTStatusID, intGPSStatusID, strIMEI, intTamperTag);
                    }

                    TheVehicleClass.UpdateOilChangeInformation(intVehicleID, intOdometer, DateTime.Now);
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Vehicle Info // Main Window // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

            PleaseWait.Close();
        }
    }
}
