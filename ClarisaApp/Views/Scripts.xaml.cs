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
using System.Collections;

namespace ClarisaApp.Views
{
    /// <summary>
    /// Interaction logic for Scripts.xaml
    /// </summary>
    public partial class Scripts : RadWindow
    {
        public Scripts(decimal idEjecutor,decimal idPOM)
        {
            InitializeComponent();
            chkHabilitar.IsChecked= true;
            chkDeshabilitar.IsChecked = false;
            lblPOM.Content = idPOM;
            lblTitulo.Content = idEjecutor;

        }

        private void RadRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            lstScripts.SelectAll();
            chkDeshabilitar.IsChecked = false;
        }

        private void RadRadioButton_Checked_1(object sender, RoutedEventArgs e)
        {
            lstScripts.UnselectAll();
            chkHabilitar.IsChecked = false;
        }

        private void RadButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                Datos datos = new DAL.Datos();
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


                IList lista = lstScripts.SelectedItems;



                //INSERTA SI NO EXISTE EJECUTOR CALENDARIZADO
                DataTable dtEjecutor = datos.VerificaEjecutorCalendarizado(Convert.ToDecimal(lblTitulo.Content), sNombre).Tables[0];
                if (dtEjecutor.Rows.Count != 0)
                {
                    //ACTUALIZA EJECUTOR
                    foreach (var item in lstScripts.SelectedItems)
                    {

                       var a=  item.ToString();
                       
                      
                        datos.CalendarizarEjecutor1(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor2(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor3(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor4(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor5(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor6(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor7(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor8(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor9(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor10(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor11(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor12(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor13(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor14(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor15(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor16(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor17(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);

                        //Politicas de Pago
                        datos.CalendarizarPolitica180(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarPolitica120(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarPolitica90(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarPolitica60(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarPolitica45(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarPolitica2030(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);



                        datos.CalendarizarEjecutor18(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor19(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor20(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor21(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor22(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor23(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor24(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor25(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                        datos.CalendarizarEjecutor26(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);

                    }
                }
                else
                {
                    //INSERTA EJECUTOR
                    datos.CalendarizarEjecutor(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);

                    datos.CalendarizarEjecutor1(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor2(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor3(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor4(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor5(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor6(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor7(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor8(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor9(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor10(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor11(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor12(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor13(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor14(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor15(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor16(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor17(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);


                    //Politicas de Pago
                    datos.CalendarizarPolitica180(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarPolitica120(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarPolitica90(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarPolitica60(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarPolitica45(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarPolitica2030(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);


                    datos.CalendarizarEjecutor18(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor19(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor20(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor21(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor22(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor23(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor24(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor25(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);
                    datos.CalendarizarEjecutor26(Convert.ToDecimal(lblTitulo.Content), Convert.ToDecimal(lblPOM.Content), FechaInicio, FechaFin, sNombre);




                }



                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "El Ejecutor se calendarizó con exito.";
                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);

            }
            catch(Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex.Message;
                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }
        }
    }
}
