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
using Telerik.Windows.Controls;

namespace ClarisaApp.Views
{
    /// <summary>
    /// Lógica de interacción para POM.xaml
    /// </summary>
    public partial class POM : Page
    {
       
       
        

        public POM(decimal idPOM)
        {
           

            InitializeComponent();
            AddTreeViewItems(idPOM);
            lblPOM.Content = idPOM;
        }

        private void AddTreeViewItems(decimal idPOM)
        {
            //Llenna arbol del proceso


            DAL.Datos datos = new DAL.Datos();
  
            IEnumerable<DataRow> procesoQuery =
             from proceso in datos.ObtenerProceso(idPOM).Tables[0].AsEnumerable()
            select proceso;


            RadTreeViewItem POM = new RadTreeViewItem();

            IEnumerable<DataRow> CatPOM =
            procesoQuery.Where(p => p.Field<string>("Categoria") == "POM");

            foreach (DataRow proceso in CatPOM)
            {
                POM.Header = proceso.Field<string>("Nombre");
                POM.Uid = Convert.ToString(proceso.Field<Int32>("idCatPOM_Ejecutores"));
                POM.IsExpanded = true;
                radTreeViewPOM.Items.Add(POM);

              
                IEnumerable<DataRow> CatCat =
               procesoQuery.Where(p => p.Field<Int32>("Padre") == proceso.Field<Int32>("Id"));

                foreach (DataRow cat in CatCat)
                {
                    RadTreeViewItem category = new RadTreeViewItem();
                    category.Header = cat.Field<string>("Nombre");
                    category.IsExpanded = true;
                    category.Uid = Convert.ToString(cat.Field<Int32>("idCatPOM_Ejecutores"));
                    category.Cursor = Cursors.Hand;
                    if(cat.Field<string>("Nombre")=="Ejecutores")
                    {
                        category.ToolTip = "Da click para agregar mas ejecutores.";
                    }
                    
                    // category.Foreground = new SolidColorBrush(Colors.Green);
                    POM.Items.Add(category);


                    IEnumerable<DataRow> CatHijo =
                    procesoQuery.Where(p => p.Field<Int32>("Padre") == cat.Field<Int32>("Id"));

                    foreach (DataRow hijo in CatHijo)
                    {
                        RadTreeViewItem product = new RadTreeViewItem();
                        product.Header = hijo.Field<string>("Nombre");
                        product.Uid = Convert.ToString(hijo.Field<Int32>("idCatPOM_Ejecutores"));
                        category.Items.Add(product);
                    }

                }


            }


        }


        

   


        private void radTreeViewPOM_ItemClick(object sender, Telerik.Windows.RadRoutedEventArgs e)
        {
            

            object item = radTreeViewPOM.SelectedValue;
            if (item != null)
            {
                RadTreeViewItem tvi = radTreeViewPOM.ContainerFromItemRecursive(radTreeViewPOM.SelectedValue);
                if (tvi != null)
                {

                    if (Convert.ToString(item) == "Ejecutores")
                    {
                        Nuevo_Ejecutor NewPage = new Nuevo_Ejecutor();
                        NewPage.MyVar = Convert.ToInt32(tvi.Uid);
                        MainFramePOM.NavigationService.Navigate(NewPage);
                        lblTitulo.Content = "Agregar nuevo ejecutor";
                    }
                    else if (Convert.ToString(item) == "Requerido")
                    {
                        MainFramePOM.NavigationService.Navigate(new Requerido(Convert.ToDecimal(lblPOM.Content)));
                        lblTitulo.Content = "Requerido";
                    }
                    else if(Convert.ToString(item) =="Primer Ajuste")
                    {
                        MainFramePOM.NavigationService.Navigate(new Ajuste(Convert.ToDecimal(lblPOM.Content)));
                        lblTitulo.Content = "Ajuste por pozo";
                    }
                    else
                    {
                        DAL.Datos datos = new DAL.Datos();

                        DataTable dt = datos.ObtenerNombreEjecutor(Convert.ToDecimal(tvi.Uid)).Tables[0];

                        foreach(DataRow row in dt.Rows)
                        {
                            lblTitulo.Content = row["Nombre"].ToString();
                        }

                        MainFramePOM.NavigationService.Navigate(new Ejecutores(Convert.ToDecimal(tvi.Uid),Convert.ToDecimal(lblPOM.Content)));

                        //Ejecutores NewPage = new Ejecutores();
                        //NewPage.MyVar = Convert.ToInt32(tvi.Uid);

                        //  MainFramePOM.NavigationService.Navigate(NewPage);
                    }
              
                }


            }

        }

     


    }
}
