using ClarisaApp.DAL;
using Microsoft.Win32;
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


namespace ClarisaApp.Views
{
    /// <summary>
    /// Lógica de interacción para BuscarPOM.xaml
    /// </summary>
    public partial class BuscarPOM : RadWindow
    {
        public BuscarPOM()
        {
            InitializeComponent();
            lblTitulo.Content = "Para visualizar datos tendrán que estar en el directorio " + AppDomain.CurrentDomain.BaseDirectory;
            BuscaPOM();
        }


 

        public void BuscaPOM()
        {
            /* string[] ubicacion = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.accdb");

             for (int i = 0; i < ubicacion.Length; i++)
             {
                 listPOM.Items.Add(System.IO.Path.GetFileName(ubicacion[i]));

             }*/

            var con = new OleDbConnection();
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";
           

            string[] restrictions = new string[4];
            restrictions[3] = "Table";
            con.Open();
            DataTable tabls = con.GetSchema("Tables", restrictions);

            DataRow[] result = tabls.Select("TABLE_NAME LIKE 'EjecutoresCalendarizados*'");
            foreach (DataRow row in result)
            {
                listPOM.Items.Add(row["TABLE_NAME"]);
            }

            con.Close();
        

        }



        private void RadButton_Click(object sender, RoutedEventArgs e)
        {
            BuscaPOM();
        }

        private void RadButton_Click_1(object sender, RoutedEventArgs e)
        {

            //this.Closed += new EventHandler<WindowClosedEventArgs>(window_Closed);

            Datos datos = new Datos();

            var padre = MainWindow.GetWindow(new MainWindow()) as MainWindow;

            var item = listPOM.SelectedItem;

        
                DataTable dtPOM = datos.ObtenerIdPOM(item.ToString()).Tables[0];
                decimal idPOM = 0;
                foreach (DataRow row in dtPOM.Rows)
                {
                    idPOM = Convert.ToDecimal(row["Id"]);
                }

                padre.MainFrame.NavigationService.Navigate(new POM(idPOM));
                padre.Show();
          

            this.Close();
        }

        //private void window_Closed(object sender, WindowClosedEventArgs e)
        //{
           

        //}
    }
}
