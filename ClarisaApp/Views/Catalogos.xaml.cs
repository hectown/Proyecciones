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

namespace ClarisaApp.Views
{
    /// <summary>
    /// Interaction logic for Catalogos.xaml
    /// </summary>
    public partial class Catalogos : RadWindow
    {
        OleDbConnection con;
        DataTable dt;
       

        public Catalogos()
        {
            InitializeComponent();
            cmbCatalogos.Text = "Actividades";
            btGuardar.Visibility = Visibility.Hidden;
            btCancelar.Visibility = Visibility.Hidden;
            //    con = new OleDbConnection();
            //  con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDetallePot2Cantarell2017.accdb";


            // BindGrid();



        }


        private void BindGrid()
        {
            /* OleDbCommand cmd = new OleDbCommand();
             if (con.State != ConnectionState.Open)
                 con.Open();
             cmd.Connection = con;
             cmd.CommandText = "Select * from CatProyectos";
             OleDbDataAdapter da = new OleDbDataAdapter(cmd);
             dt = new DataTable();
             da.Fill(dt);
             gvData.ItemsSource = dt.AsDataView();*/

          //  gvData.ItemsSource = Datos.ObtenerCatalogos("Contratos").Tables[0].AsDataView();

         

        }

        private void RadComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lblTItulo.Content = "Catálogo " + cmbCatalogos.SelectedItem;
            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += new DoWorkEventHandler(llenarCombo);
            bw.ProgressChanged += new ProgressChangedEventHandler(llenarComboChanged);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(LLenarComboFin);
         
            bw.RunWorkerAsync();
         
        }

        void llenarCombo(object sender, DoWorkEventArgs e)
        {

            //Aqui el codigo o la llamada a la funcion que tarda en ejecutarse

            Datos dt = new Datos();
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                try
                {
                    gvData.ItemsSource = dt.ObtenerCatalogos(cmbCatalogos.Text).Tables[0].AsDataView();
                }
                catch(Exception ex)
                {
                   
                }
               
            }), null);
        
            //NOTA: No puedes interactuar con la interfaz grafica debido a que se ejecuta en otro hilo
        }

        void llenarComboChanged(object sender, ProgressChangedEventArgs e)
        {
            //Aqui el codigo para mostrar progreso
            
            
            //NOTA: Se puede usar la interfaz grafica
        }

        void LLenarComboFin(object sender, RunWorkerCompletedEventArgs e)
        {
            //Aqui el codigo a ejecutar cuando finalize la ejecucion
          
            //NOTA: Se puede usar la interfaz grafica
        }




        private void radGridView_CellValidating(object sender, GridViewCellValidatingEventArgs e)
        {
            if (e.Cell.Column.UniqueName == "POSPRE")
            {
                string newValue = e.NewValue.ToString();
                if (String.IsNullOrEmpty(newValue))
                {
                    e.IsValid = false;
                    e.ErrorMessage = "No debe de ser nulo";
                }
                else
                {
                    e.IsValid = false;
                    e.ErrorMessage = "No debe de ser nulo";  
                        

                }
                      
            }
        }
    
        private void RadButton_Click(object sender, RoutedEventArgs e)
        {


            btGuardar.Visibility = Visibility.Visible;
            btCancelar.Visibility = Visibility.Visible;

            radBusyIndicator.IsBusy = true;

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
            //var padre = Window.GetWindow(this) as MainWindow;
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
            //var padre = Window.GetWindow(this) as MainWindow;
            //padre.radBusyIndicator.IsBusy = true;
            radBusyIndicator.IsBusy = true;

            MessageBoxResult messageBoxResult = System.Windows.MessageBox.Show("¿Estas seguro de modificar el catálogo?", "Atención", System.Windows.MessageBoxButton.YesNo);
            if (messageBoxResult == MessageBoxResult.Yes)
            {


               

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(actualizarCatalogo);
                bw.ProgressChanged += new ProgressChangedEventHandler(actualizarCatalogoChanged);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(actualizarCatalogoFin);

                bw.RunWorkerAsync();


            }
            else
            {
                radBusyIndicator.IsBusy = false;
                //padre.radBusyIndicator.IsBusy = false;
            }

        }

        void actualizarCatalogo(object sender, DoWorkEventArgs e)
        {

            //Aqui el codigo o la llamada a la funcion que tarda en ejecutarse


            Datos dt = new Datos();
            this.Dispatcher.Invoke(new System.Action(() =>
            {

 
                dt.BorrarCatalogos(cmbCatalogos.Text);
                dt.GuardarCatActividades(gvData, cmbCatalogos.Text);

            }), null);

            //NOTA: No puedes interactuar con la interfaz grafica debido a que se ejecuta en otro hilo
        }

        void actualizarCatalogoChanged(object sender, ProgressChangedEventArgs e)
        {
            //Aqui el codigo para mostrar progreso


            //NOTA: Se puede usar la interfaz grafica
        }

        void actualizarCatalogoFin(object sender, RunWorkerCompletedEventArgs e)
        {
            //var padre = Window.GetWindow(this) as MainWindow;
            //Aqui el codigo a ejecutar cuando finalize la ejecucion
            Datos dt = new Datos();
            gvData.ItemsSource = dt.ObtenerCatalogos(cmbCatalogos.Text).Tables[0].AsDataView();
            //padre.radBusyIndicator.IsBusy = false;
            radBusyIndicator.IsBusy = false;
            btGuardar.Visibility = Visibility.Hidden;
            btCancelar.Visibility = Visibility.Hidden;

            //NOTA: Se puede usar la interfaz grafica
        }

        private void RadButtonCancelar_Click(object sender, RoutedEventArgs e)
        {
            Datos dt = new Datos();
            gvData.ItemsSource = dt.ObtenerCatalogos(cmbCatalogos.Text).Tables[0].AsDataView();
            btGuardar.Visibility = Visibility.Hidden;
            btCancelar.Visibility = Visibility.Hidden;
        }
    }
}
