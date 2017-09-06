using System;
using System.Collections.Generic;
using System.Data;
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
using ClarisaApp.DAL;

namespace ClarisaApp.Views
{
    /// <summary>
    /// Interaction logic for Ajuste.xaml
    /// </summary>
    public partial class Ajuste : Page
    {
        public Ajuste(decimal idPOM)
        {
            InitializeComponent();
            ObtenerAjuste(idPOM);
            lblPOM.Content = idPOM;
        }


        public void ObtenerAjuste(decimal idPOM)
        {
            DAL.Datos datos = new DAL.Datos();
            DataTable dt = datos.ObtenerAñosPOM(idPOM).Tables[0];
            decimal FechaInicio = 0;
            decimal FechaFin = 0;
            string sNombre = "";

            foreach (DataRow row in dt.Rows)
            {
                sNombre = "Ajuste" + row["Nombre"].ToString();
                FechaInicio = Convert.ToDecimal(row["Fecha_Inicio"]);
                FechaFin = Convert.ToDecimal(row["Fecha_Fin"]);
            }

            if (datos.ObtenerRequerido(sNombre).Tables[0].Rows.Count == 0)
            {
                lblTitulo.Content = "Actualmente no se ha realizado ningun ajuste.";
                gvData.Visibility = Visibility.Hidden;
                gvDataPager.Visibility = Visibility.Hidden;

            }
            else
            {
                gvData.ItemsSource = datos.ObtenerRequerido(sNombre).Tables[0];
            }




        }

        private void btAjustar_Click(object sender, RoutedEventArgs e)
        {
            Datos datos = new DAL.Datos();
            try
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

                DataTable dtAjustePozo = datos.CrearTablaAjusteporPozo(Convert.ToDecimal(lblPOM.Content),sNombre,FechaInicio,FechaFin).Tables[0];
                




            }
            catch(Exception ex)
            {

            }

        }

        private void btCostoPozo_Click(object sender, RoutedEventArgs e)
        {
            
              new CostoPozos(Convert.ToDecimal(lblPOM.Content)).Show();
        }

    }
}
