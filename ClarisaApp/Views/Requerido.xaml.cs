using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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
using Telerik.Windows.Documents.Core;
using Telerik.Windows.Documents.Fixed;
using Telerik.Windows.Documents.Spreadsheet;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.Pdf;
using Telerik.Windows.Zip;



namespace ClarisaApp.Views
{
    /// <summary>
    /// Interaction logic for Requerido.xaml
    /// </summary>
    public partial class Requerido : Page
    {
        public Requerido(decimal idPOM)
        {
            InitializeComponent();
            ObtenerRequerido(idPOM);

        }

       public void ObtenerRequerido(decimal idPOM)
        {
            DAL.Datos datos = new DAL.Datos();
            DataTable dt= datos.ObtenerAñosPOM(idPOM).Tables[0];
            decimal FechaInicio = 0;
            decimal FechaFin = 0;
            string sNombre = "";

            foreach (DataRow row in dt.Rows)
            {
                sNombre = row["Nombre"].ToString();
                FechaInicio = Convert.ToDecimal(row["Fecha_Inicio"]);
                FechaFin = Convert.ToDecimal(row["Fecha_Fin"]);
            }

            if (datos.ObtenerRequerido(sNombre).Tables[0].Rows.Count == 0)
            {
                lblTitulo.Content = "Actualmente no hay ningun ejecutor calendarizado.";
                gvData.Visibility= Visibility.Hidden;
                gvDataPager.Visibility = Visibility.Hidden;

            }else
            {
                gvData.ItemsSource = datos.ObtenerRequerido(sNombre).Tables[0];
            }
           
           


        }


        private void btnExportAsync_Click(object sender, EventArgs e)
        {
            string extension = "xls";
            SaveFileDialog dialog = new SaveFileDialog()
            {
                DefaultExt = extension,
                Filter = String.Format("{1} files (*.{0})|*.{0}|All files (*.*)|*.*", extension, "Excel"),
                FilterIndex = 1
            };
            if (dialog.ShowDialog() == true)
            {
                using (Stream stream = dialog.OpenFile())
                {
                    gvData.Export(stream,
                     new GridViewExportOptions()
                     {
                         Format = ExportFormat.Html,
                         ShowColumnHeaders = true,
                         ShowColumnFooters = true,
                         ShowGroupFooters = false,
                     });
                }
            }
        }

    }
}
