using ClarisaApp.DAL;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using Telerik.Windows.Controls;

namespace ClarisaApp.Views
{
    /// <summary>
    /// Interaction logic for CostoPozos.xaml
    /// </summary>
    public partial class CostoPozos
    {
        public CostoPozos(decimal idPOM)
        {
            InitializeComponent();
            inicio(idPOM);
            lblPOM.Content = idPOM;
        }

        void inicio(decimal idPOM)
        {

            radBusyIndicator.IsBusy = true;
            DAL.Datos datos = new DAL.Datos();
            DataTable dt = datos.ObtenerAñosPOM(idPOM).Tables[0];
            decimal FechaInicio = 0;
            decimal FechaFin = 0;
            string sNombre = "";

            foreach (DataRow row in dt.Rows)
            {
                sNombre = row["Nombre"].ToString();
                FechaInicio = Convert.ToDecimal(row["Fecha_Inicio"]);
                FechaFin = Convert.ToDecimal(row["Fecha_Fin"]);
            }
            
            string nombre = sNombre.Substring(24);

            if (datos.ObtenerCostoPozo(nombre).Tables[0].Rows.Count == 0)
            {
                

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(importarExcel);
                bw.ProgressChanged += new ProgressChangedEventHandler(importarExcelChanged);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(importarExcelFin);

                bw.RunWorkerAsync();
            }
            else
            {
                

       

                gvData.ItemsSource = datos.ObtenerCostoPozo(nombre).Tables[0];

         

               

               
                radBusyIndicator.IsBusy = false;
            }


        }

        private void btCargarCostoPozo_Click(object sender, RoutedEventArgs e)
        {
            var padre = Window.GetWindow(this) as MainWindow;

          

         //   padre.radBusyIndicator.IsBusy = true;

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
            
            //Aqui el codigo a ejecutar cuando finalize la ejecucion
            //padre.radBusyIndicator.IsBusy = false;
            radBusyIndicator.IsBusy = false;

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

                    //string strCellData = "";
                    //double douCellData;
                    int rowCnt = 0;
                    int colCnt = 0;
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("es-MX");
                    DataTable dt = new DataTable();
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        string strColumn = "";
                        strColumn = (string)(excelRange.Cells[1, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        dt.Columns.Add(strColumn, typeof(string));
                    }

                    int columns_count = excelRange.Columns.Count;
                    int rows_count = excelRange.Rows.Count;

                    for (rowCnt = 2; rowCnt <= rows_count; rowCnt++)
                    {
                        object[] strData = new object[columns_count];

                        for (int clm = 0; clm < columns_count; clm++)
                        {
                            strData[clm] = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[rowCnt, clm + 1]).Value2;
                        }


                        dt.Rows.Add(strData);
                  

                    }

                    gvData.ItemsSource = dt.DefaultView;

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
                   
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");

            }



        }



        private void btGuardarCostoPozo_Click(object sender, RoutedEventArgs e)
        {
            radBusyIndicator.IsBusy = true;

            MessageBoxResult messageBoxResult = System.Windows.MessageBox.Show("¿Estas seguro de guardar?", "Atención", System.Windows.MessageBoxButton.YesNo);
            if (messageBoxResult == MessageBoxResult.Yes)
            {




                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(actualizarCosto);
                bw.ProgressChanged += new ProgressChangedEventHandler(actualizarCostoChanged);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(actualizarCostoFin);

                bw.RunWorkerAsync();


            }
            else
            {
                radBusyIndicator.IsBusy = false;
                //padre.radBusyIndicator.IsBusy = false;
            }
        }


        void actualizarCosto(object sender, DoWorkEventArgs e)
        {

            //Aqui el codigo o la llamada a la funcion que tarda en ejecutarse


            Datos datos = new Datos();
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                
                DataTable dt = datos.ObtenerAñosPOM(Convert.ToDecimal(lblPOM.Content)).Tables[0];
                decimal FechaInicio = 0;
                decimal FechaFin = 0;
                string sNombre = "";

                foreach (DataRow row in dt.Rows)
                {
                    sNombre = row["Nombre"].ToString();
                    FechaInicio = Convert.ToDecimal(row["Fecha_Inicio"]);
                    FechaFin = Convert.ToDecimal(row["Fecha_Fin"]);
                }

                string nombre = sNombre.Substring(24);

                datos.BorrarCosto(nombre);
                datos.GuardarCosto(gvData, nombre);

            }), null);

            //NOTA: No puedes interactuar con la interfaz grafica debido a que se ejecuta en otro hilo
        }

        void actualizarCostoChanged(object sender, ProgressChangedEventArgs e)
        {
            //Aqui el codigo para mostrar progreso


            //NOTA: Se puede usar la interfaz grafica
        }

        void actualizarCostoFin(object sender, RunWorkerCompletedEventArgs e)
        {
            
            //Aqui el codigo a ejecutar cuando finalize la ejecucion
            Datos datos = new Datos();

            DataTable dt = datos.ObtenerAñosPOM(Convert.ToDecimal(lblPOM.Content)).Tables[0];
            decimal FechaInicio = 0;
            decimal FechaFin = 0;
            string sNombre = "";

            foreach (DataRow row in dt.Rows)
            {
                sNombre = row["Nombre"].ToString();
                FechaInicio = Convert.ToDecimal(row["Fecha_Inicio"]);
                FechaFin = Convert.ToDecimal(row["Fecha_Fin"]);
            }

            string nombre = sNombre.Substring(24);

            gvData.ItemsSource = datos.ObtenerCostoPozo(nombre).Tables[0].AsDataView();
            radBusyIndicator.IsBusy = false;
            radBusyIndicator.IsBusy = false;
            btGuardar.Visibility = Visibility.Hidden;
        

            //NOTA: Se puede usar la interfaz grafica
        }

        private void gvData_AutoGeneratingColumn(object sender, GridViewAutoGeneratingColumnEventArgs e)
        {
            GridViewDataColumn column = e.Column as GridViewDataColumn;
            // set the cell format of numeric values.
            if (column.DataType.Name != "TOTAL")
            {
                column.DataFormatString = "###,###";
                column.TextAlignment = TextAlignment.Right;
            }
        }
    }
}
