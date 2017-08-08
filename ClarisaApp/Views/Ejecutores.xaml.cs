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
            btCancelar.Visibility = Visibility.Hidden;

        }

       

   


   
      


  
       
        public void llenarGridEjecutores(decimal parametros)
        {
            var a = parametros;
            Datos datos = new Datos();



            gvData.ItemsSource = datos.ObtenerEjecutor(parametros).Tables[0]; ;

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
    }
}
