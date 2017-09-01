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
    /// Lógica de interacción para Ejecutores.xaml
    /// </summary>
    public partial class Ejecutores : Page
    {

        private Int32 pMyVar;


        public Int32 MyVar
        {
            get { return this.pMyVar; }
            set { this.pMyVar = value; }
        }


        public Ejecutores(decimal idEjecutor, decimal idPOM)
        {
            InitializeComponent();
            llenarGridEjecutores(idEjecutor);
            lblTitulo.Content = idEjecutor;
            lblPOM.Content = idPOM;
         

        }



        public void llenarGridEjecutores(decimal parametros)
        {
            var a = parametros;
            Datos datos = new Datos();



            gvData.ItemsSource = datos.ObtenerEjecutor(parametros).Tables[0]; 
           
          

        }

        private void btCalendarizar_Click(object sender, RoutedEventArgs e)
        {
            
            try
            {
                

                new Scripts(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content)).Show();

            }
            catch (Exception ex)
            {

            }

        }

        private void btBorrar_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult messageBoxResult = System.Windows.MessageBox.Show("¿Estas seguro de borrar el ejecutor?", "Atención", System.Windows.MessageBoxButton.YesNo);
            if (messageBoxResult == MessageBoxResult.Yes)
            {

                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(BorrarEjecutor);
                bw.ProgressChanged += new ProgressChangedEventHandler(BorrarEjecutorChanged);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BorrarEjecutorFin);

                bw.RunWorkerAsync();





            }
            else
            {
            }
        }



        void BorrarEjecutor(object sender, DoWorkEventArgs e)
        {

            this.Dispatcher.Invoke(new System.Action(() =>
            {

                Datos datos = new DAL.Datos();
            DataTable dt = datos.ObtenerAñosPOM(Convert.ToDecimal(lblPOM.Content)).Tables[0];
            decimal FechaInicio = 0;
            decimal FechaFin = 0;
            string sNombre = "";
            decimal idEjecutor = Convert.ToDecimal(lblTitulo.Content);

            foreach (DataRow row in dt.Rows)
            {
                sNombre = row["Nombre"].ToString();
                FechaInicio = Convert.ToDecimal(row["Fecha_Inicio"]);
                FechaFin = Convert.ToDecimal(row["Fecha_Fin"]);
            }

         

                var a = datos.BorrarEjecutor(Convert.ToDecimal(lblPOM.Content), idEjecutor, sNombre);
            if (a == true)
            {
                var b = datos.BorrarEjecutoresTabla(Convert.ToDecimal(lblPOM.Content), idEjecutor);

                if (b == true)
                {
                    var c = datos.BorrarEjecutoresEstructura(Convert.ToDecimal(lblPOM.Content), idEjecutor);
                }
            }
            else
            {
            }
            }), null);
        }



        void BorrarEjecutorChanged(object sender, ProgressChangedEventArgs e)
        {
            //Aqui el codigo para mostrar progreso


            //NOTA: Se puede usar la interfaz grafica
        }

        void BorrarEjecutorFin(object sender, RunWorkerCompletedEventArgs e)
        {
            var padre = Window.GetWindow(this) as MainWindow;
            //Aqui el codigo a ejecutar cuando finalize la ejecucion
            //Datos dt = new Datos();
            //gvData.ItemsSource = dt.ObtenerCatalogos(cmbCatalogos.Text).Tables[0].AsDataView();
            padre.radBusyIndicator.IsBusy = false;
           

            padre.MainFrame.NavigationService.Navigate(new POM(Convert.ToDecimal(lblPOM.Content)));

            //NOTA: Se puede usar la interfaz grafica
        }

       
      


    }
}
