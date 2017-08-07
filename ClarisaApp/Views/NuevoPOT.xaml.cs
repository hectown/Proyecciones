using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
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
using Telerik.Windows.Controls;
using ClarisaApp.DAL;

namespace ClarisaApp.Views
{
    /// <summary>
    /// Interaction logic for NuevoPOT.xaml
    /// </summary>
    public partial class NuevoPOT : RadWindow
    {
        public NuevoPOT()
        {
            InitializeComponent();
            this.dpFecha.Culture = new System.Globalization.CultureInfo("es-MX");
            this.dpFecha.Culture.DateTimeFormat.ShortDatePattern = "yyyy";
            dpFecha.SelectedDate = DateTime.Now;
        }

        private void dpFecha_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
           
        }

        public void cambiaFechas()
        {
            lblFechas.Content = "El POT se creará desde " + dpFecha.SelectedValue +" a "+ dpFecha.SelectedValue + 1;
        }

        private void RadButton_Click(object sender, RoutedEventArgs e)
        {
            Datos dat = new DAL.Datos();

            var inicio = dpFecha.SelectedDate.Value.Year;
            var fin = inicio + 1;
            var nombrePOT = "EjecutoresCalendarizados" + txtNombre.Text;
            var nombrePOTAjuste ="Ajuste"+ nombrePOT;

            dat.GuardarPOM(nombrePOT,inicio.ToString(),fin.ToString());
            var a = dat.CrearNuevoPOT(nombrePOT, inicio.ToString(),fin.ToString());

            if (a == true)
                {
                dat.CrearNuevoPOT(nombrePOTAjuste, inicio.ToString(), fin.ToString());
                this.Close();
            DataTable dtPOM = dat.ObtenerIdPOM(nombrePOT).Tables[0];
            decimal idPOM = 0;
            foreach (DataRow row in dtPOM.Rows)
            {
                idPOM = Convert.ToDecimal(row["Id"]);
            }

              

            var padre = MainWindow.GetWindow(new MainWindow()) as MainWindow;
            padre.MainFrame.NavigationService.Navigate(new POM(idPOM));
            padre.Show();

                
            }
        }
    }
}
