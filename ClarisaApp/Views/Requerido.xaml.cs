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
using Telerik.Windows.Controls.GridView;
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
        string nameBD;
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
                nameBD = sNombre;
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

        private void gvData_RowLoaded(object sender, Telerik.Windows.Controls.GridView.RowLoadedEventArgs e)
        {
            int countC = e.Row.Cells.Count;
            for (int i = 0; i < countC; i++)
            {

                var _cell = e.Row.Cells[i];
                if (_cell.DataContext.ToString() == "System.Data.DataRow")
                {
                    GridViewCell _cellValid = (GridViewCell)e.Row.Cells[i];

                    if (_cellValid.Value == null)
                    {
                        var converter = new System.Windows.Media.BrushConverter();
                        var brush = (Brush)converter.ConvertFromString("#D32F2F");
                        e.Row.Cells[i].Background = brush;
                    }
                }


            }
        }

        private void gvData_CellEditEnded(object sender, GridViewCellEditEndedEventArgs e)
        {

            DAL.Datos datos = new DAL.Datos();
            GridViewCell cl = (GridViewCell)e.Cell;
            TextBox t = (TextBox)e.EditingElement;
            Int32 c = e.Cell.TabIndex;

            GridViewCell _id= (GridViewCell)e.Cell.ParentRow.Cells[0];

            RadGridView _gv = (RadGridView)e.Source;

            string _columnHeader = _gv.CurrentColumn.Header.ToString();
            datos.UpdateRequerido(nameBD, _id.Value.ToString(),  _columnHeader, t.Text);

            //gvData.Rebind();
            //datos.ObtenerRequerido(sNombre).Tables[0];

            //MessageBox.Show(_id.Value.ToString());
        }
    }
}
