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
using Telerik.Windows.Controls;
using ClarisaApp.Views;
using System.Globalization;
using System.Threading;

namespace ClarisaApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        OleDbConnection con;
        DataTable dt;
        public MainWindow()
        {
            InitializeComponent();

          

        }

        public RadBusyIndicator radOcupado
        {
            get
            {
                return radBusyIndicator;
            }
        }

        private void Catalogos_Click(object sender, RoutedEventArgs e)
        {

            new Catalogos().Show();
            //MainFrame.Source = new Uri("/Views/Catalogos.xaml", UriKind.Relative);
        }


        private void AplicarEfecto(Window win)

        {

            var objBlur = new System.Windows.Media.Effects.BlurEffect();

            objBlur.Radius = 5;

            win.Effect = objBlur;

        }

    

        private void ButtonNuevo_Click(object sender, RoutedEventArgs e)
        {


            // RadWindow.Prompt("Nuevo POT",this.OnClosed);

            new NuevoPOT().Show();
            //MainFrame.Source = new Uri("/Views/POM.xaml", UriKind.Relative);

        }

        private void OnClosed(object sender, WindowClosedEventArgs e)
        {
           if(e.DialogResult==true)
            {
                var textbox="EjecutoresCalendarizados"+e.PromptResult.ToString();

                DAL.Datos datos = new DAL.Datos();


                // datos.GuardarPOM(textbox);
               // datos.CrearBaseDatos(textbox);
                //datos.CrearNuevoPOT(textbox);

                //new POM(Convert.ToDecimal(4));
                //FALTA REDIERCIONAR A PAGINA DE POM
                
               /* DataTable dtPOM  = datos.ObtenerIdPOM(textbox).Tables[0];
                decimal idPOM = 0;
                foreach (DataRow row in dtPOM.Rows)
                {
                    idPOM = Convert.ToDecimal(row["Id"]);
                }

                MainFrame.NavigationService.Navigate(new POM(Convert.ToDecimal(idPOM)));*/

            }




        }

        private void AbrirPOM_Click(object sender, RoutedEventArgs e)
        {
          
            new BuscarPOM().Show();
            

        }
    }
}
