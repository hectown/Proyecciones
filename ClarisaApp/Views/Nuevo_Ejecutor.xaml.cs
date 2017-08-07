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
using System.Data;
using System.Data.OleDb;
using ClarisaApp.DAL;
using Telerik.Windows.Controls;
using Microsoft.Office;
using Microsoft.Win32;
using System.Threading;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ClarisaApp.Views
{
    /// <summary>
    /// Lógica de interacción para Ejecutores.xaml
    /// </summary>
    public partial class Nuevo_Ejecutor : Page
    {
        private Int32 pMyVar;
        

        public Int32 MyVar
        {
            get { return this.pMyVar; }
            set { this.pMyVar = value; }
        }
        public Nuevo_Ejecutor()
        {
            InitializeComponent();
         
            btGuardar.Visibility = Visibility.Hidden;
            btCancelar.Visibility = Visibility.Hidden;
        }

      

        private void RadButton_Click(object sender, RoutedEventArgs e)
        {
           
            var padre = Window.GetWindow(this) as MainWindow;

            btGuardar.Visibility = Visibility.Visible;
            btCancelar.Visibility = Visibility.Visible;

            padre.radBusyIndicator.IsBusy = true;

            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += new DoWorkEventHandler(importarExcel);
            bw.ProgressChanged += new ProgressChangedEventHandler(importarExcelChanged);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(importarExcelFin);

            bw.RunWorkerAsync();




        }


        void importarExcel(object sender, DoWorkEventArgs e)
        {

            //Aqui el codigo o la llamada a la funcion que tarda en ejecutarse


            this.Dispatcher.Invoke(new System.Action(() =>
            {

                importa();

            }), null);

            //NOTA: No puedes interactuar con la interfaz grafica debido a que se ejecuta en otro hilo
        }

        void importarExcelChanged(object sender, ProgressChangedEventArgs e)
        {
            //Aqui el codigo para mostrar progreso


            //NOTA: Se puede usar la interfaz grafica
        }

        void importarExcelFin(object sender, RunWorkerCompletedEventArgs e)
        {
            var padre = Window.GetWindow(this) as MainWindow;
            //Aqui el codigo a ejecutar cuando finalize la ejecucion
            padre.radBusyIndicator.IsBusy = false;


            //NOTA: Se puede usar la interfaz grafica
        }


        public void importa()
        {


            try
            {
                OpenFileDialog openfile = new OpenFileDialog();
                openfile.DefaultExt = ".xlsx";
                openfile.Filter = "(.xlsx)|*.xlsx";
                //openfile.ShowDialog();

                var browsefile = openfile.ShowDialog();

                if (browsefile == true)
                {
                    txtFilePath.Text = openfile.FileName;

                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    //Static File From Base Path...........
                    //Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "TestExcel.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    //Dynamic File Using Uploader...........
                    Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
                    Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                    string strCellData = "";
                    double douCellData;
                    int rowCnt = 0;
                    int colCnt = 0;

                    DataTable dt = new DataTable();
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        string strColumn = "";
                        strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        dt.Columns.Add(strColumn, typeof(string));
                    }

                    for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                    {
                        string strData = "";
                        for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                        {
                            try
                            {
                                strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                                strData += strCellData + "|";
                            }
                            catch (Exception ex)
                            {
                                douCellData = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                                strData += douCellData.ToString() + "|";
                            }
                        }
                        strData = strData.Remove(strData.Length - 1, 1);
                        dt.Rows.Add(strData.Split('|'));
                    }

                    gvData.ItemsSource = dt;

                    excelBook.Close(true, null, null);
                    excelApp.Quit();

                    var alert = new RadDesktopAlert();
                    alert.Header = "NOTIFICACIÓN";
                    alert.Content = "El archivo se cargó exitosamente.";
                    alert.ShowDuration = 3000;
                    RadDesktopAlertManager manager = new RadDesktopAlertManager();
                    manager.ShowAlert(alert);
                }
                else
                {
                    btGuardar.Visibility = Visibility.Hidden;
                    btCancelar.Visibility = Visibility.Hidden;
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");

            }


        }



        private void RadButtonGuardar_Click(object sender, RoutedEventArgs e)
        {
           if(txtNombreEjecutor.Text=="")
            {
                txtNombreEjecutor.Background = Brushes.Orange;
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "Por favor registra un nombre al ejecutor.";

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }
            else
            {
                var padre = Window.GetWindow(this) as MainWindow;
                padre.radBusyIndicator.IsBusy = true;


                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(GuardarEjecutor);
                bw.ProgressChanged += new ProgressChangedEventHandler(GuardarEjecutorChanged);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(GuardarEjecutorFin);

                bw.RunWorkerAsync();
            }


        }

        void GuardarEjecutor(object sender, DoWorkEventArgs e)
        {

            //Aqui el codigo o la llamada a la funcion que tarda en ejecutarse
            var idPOM = pMyVar;

            Datos dt = new Datos();
            this.Dispatcher.Invoke(new System.Action(() =>
            {
           
                dt.GuardarEjecutor(gvData,idPOM,txtNombreEjecutor.Text);

            }), null);

            //NOTA: No puedes interactuar con la interfaz grafica debido a que se ejecuta en otro hilo
        }

        void GuardarEjecutorChanged(object sender, ProgressChangedEventArgs e)
        {
            //Aqui el codigo para mostrar progreso


            //NOTA: Se puede usar la interfaz grafica
        }

        void GuardarEjecutorFin(object sender, RunWorkerCompletedEventArgs e)
        {
            var padre = Window.GetWindow(this) as MainWindow;
            //Aqui el codigo a ejecutar cuando finalize la ejecucion
            //Datos dt = new Datos();
            //gvData.ItemsSource = dt.ObtenerCatalogos(cmbCatalogos.Text).Tables[0].AsDataView();
            padre.radBusyIndicator.IsBusy = false;
            btGuardar.Visibility = Visibility.Hidden;
            btCancelar.Visibility = Visibility.Hidden;
            var idPOM = pMyVar;
          
            padre.MainFrame.NavigationService.Navigate(new POM(Convert.ToDecimal(idPOM)));
          
            //NOTA: Se puede usar la interfaz grafica
        }









    }
}
