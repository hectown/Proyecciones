using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Windows.Controls;
using System.Windows;
using Telerik.Windows.Controls;
using Telerik.Windows.Controls.GridView;
using ClarisaApp.Views;
using ADOX;
using System.Collections;

namespace ClarisaApp.DAL
{
    class Datos
    {
        

      

        private DataSet fSQL(string sSQL, OleDbParameter[] opParametros)
        {
            DataSet ds = new DataSet();
            try
            {

                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";
                if (con.State != ConnectionState.Open)
                    con.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = con;
                cmd.CommandText = sSQL;
                cmd.Parameters.AddRange(opParametros);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
               
                da.Fill(ds);
                con.Close();
                return ds;
            }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;
           
                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);


                // MessageBox.Show(ex.Message, "Error");
                return null;
            }
        }


        private static object fSQLObject(string sSQL, OleDbParameter[] opParametros)
        {

            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDetallePot2Cantarell2017.accdb";

            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = con;
            cmd.CommandText = sSQL;
            cmd.Parameters.AddRange(opParametros);

            object objDato = null;
            try
            {
                con.Open();
                objDato = cmd.ExecuteScalar();
            }
            finally
            {
                con.Close();
            }

            return objDato;
        }

        #region CATALOGOS
        public DataSet ObtenerCatalogos(string sCatalogo)
        {
            string query = "Select * from Cat"+sCatalogo;

            OleDbParameter[] opParametros = { };

            return fSQL(query, opParametros);
        }


        public void BorrarCatalogos(string Tabla)
        {
            try
            {
                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();

                string q = "Delete from Cat" + Tabla;
                OleDbCommand comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }
        }

        public bool CrearNuevoPOT(string sNombre, string sFechaIni, string sFechaFin)
        {
            try
            {
            OleDbConnection myConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb");
            myConnection.Open();
            string strTemp = " Id AUTOINCREMENT primary key, BASE Text(255), CONTRATO Text(255), FecIniContrato Text(255), FecVencContrato Text(255), RESERVA Text(255), POLITICAPAGO Text(255), MOVEQUIPOS Text(255), STATUSORIGINAL Text(255), Ficha Text(255), USUARIO Text(255), Suptcia Text(255), ACTIVO Text(255), PROYECTO Text(255), PYIN Text(255), CLASIFGRAL Text(255), PROG Text(255), CLASIFACT Text(255), CLASIFDETA Text(255), EJECUTORES Text(255), FONDO Text(255), CT Text(255), DEPTO Text(255), CTEjec Text(255), DeptoEjec Text(255), NuevoCGESTOR Text(255), LLAVECONTROL Text(255), PROPREANT Text(255), PROPRE Text(255), ElementoPEP Text(255), DESCELEMPEP Text(255), RG Text(255), CO Text(255), POZO Text(255), IIP Text(255), ACTIVIDAD Text(255), ACTIVIDADESPECIFICA Text(255), NoEQUIPO Text(255), NOMBREEQUIPO Text(255), ESTRUC Text(255), RENTA Text(255), OBJETIVO Text(255), INTERVALO Text(255), PROF Text(255), Fec_Inic DateTime, Fec_Term DateTime, Días Number, NOTAS Text(255), POSPREORIG Text(255), MONEDA Number,COSTEO_15 Text(255),PIVDEV Text(255), PIVFLU Text(255),Descripción_Contrato Text(255), DocSap Text(255), Año Text(255), Cve_Campo Text(255), Id_Asignacion Text(255), DescripcionAsignacion Text(255), Cve_Asignacion Text(255), Compañia Text(255), Acreedor Text(255), DENE" + sFechaIni+ " Number, DFEB" + sFechaIni + " Number, DMAR" + sFechaIni + " Number, DABR" + sFechaIni + " Number, DMAY" + sFechaIni + " Number, DJUN" + sFechaIni + " Number, DJUL" + sFechaIni + " Number, DAGO" + sFechaIni + " Number, DSEP" + sFechaIni + " Number, DOCT" + sFechaIni + " Number, DNOV" + sFechaIni + " Number, DDIC" + sFechaIni + " Number, FENE"+sFechaIni+ " Number, FFEB" + sFechaIni + " Number, FMAR" + sFechaIni + " Number, FABR" + sFechaIni + " Number, FMAY" + sFechaIni + " Number, FJUN" + sFechaIni + " Number, FJUL" + sFechaIni + " Number, FAGO" + sFechaIni + " Number, FSEP" + sFechaIni + " Number, FOCT" + sFechaIni + " Number, FNOV" + sFechaIni + " Number, FDIC" + sFechaIni + " Number, FENE" + sFechaFin + " Number, FFEB" + sFechaFin + " Number, FMAR" + sFechaFin + " Number, FABR" + sFechaFin + " Number, FMAY" + sFechaFin + " Number, FJUN" + sFechaFin + " Number, FJUL" + sFechaFin + " Number, FAGO" + sFechaFin + " Number, FSEP" + sFechaFin + " Number, FOCT" + sFechaFin + " Number, FNOV" + sFechaFin + " Number, FDIC" + sFechaFin + " Number, TOTDev"+sFechaIni+ " Number, TOTFe"+sFechaIni+ " Number, TOTFe" + sFechaFin + " Number, TOTFe Number, idEjecutores Number, idPOM Number";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = "CREATE TABLE " + sNombre + "(" + strTemp + ")";
            myCommand.ExecuteNonQuery();
            myCommand.Connection.Close();
                return true;
             }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex.Message;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
                return false;
            }
        }


        //Costo pozos
        public bool CrearNuevoCostoPozos(string sNombre)
        {
            try
            {
                OleDbConnection myConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb");
                myConnection.Open();
                string strTemp = " Id AUTOINCREMENT primary key, EQUIPO Text(255), PLATAFORMA Text(255), POZO Text(255), TOTAL Number";
                OleDbCommand myCommand = new OleDbCommand();
                myCommand.Connection = myConnection;
                myCommand.CommandText = "CREATE TABLE " + sNombre + "(" + strTemp + ")";
                myCommand.ExecuteNonQuery();
                myCommand.Connection.Close();
                return true;
            }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex.Message;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
                return false;
            }
        }

        public void CrearBaseDatos(string sNombre)
        {
            
            try
            {
                var catalog = new Catalog();
             
                catalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;" +
                "Data Source =" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDatos" + sNombre + ".accdb;" +
                "Jet OLEDB:Engine Type=5");
                // Create a table and two columns
                Table dt = new Table();
                dt.Name = "MyTable";
                dt.Columns.Append("col1", DataTypeEnum.adInteger, 4);
                dt.Columns.Append("col2", DataTypeEnum.adVarWChar, 255);
                //Add table to the tables collection
                catalog.Tables.Append((object)dt);
                Table dt2 = new Table();
                dt2.Name = "Tabla";
                dt2.Columns.Append("col1", DataTypeEnum.adInteger, 4);
                catalog.Tables.Append((object)dt2);
            }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex.Message;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }
        }

        public void GuardarPOM(string sNombre, string sFechaInicio, string sFechaFin)
        {
            try
            {
                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();
                string q = "INSERT INTO CatPOM (Nombre,Fecha,Fecha_Inicio,Fecha_Fin) VALUES (@Nombre,@Fecha,@FechaInicio,@FechaFin)";
                OleDbCommand comando = new OleDbCommand(q, con);

                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@Nombre", sNombre);
                comando.Parameters.AddWithValue("@Fecha", (DateTime.Now).ToString("dd/MM/yyyy"));
                comando.Parameters.AddWithValue("@FechaInicio",sFechaInicio);
                comando.Parameters.AddWithValue("@FechaFin", sFechaFin);
                comando.ExecuteNonQuery();

            con.Close();
           
                //SI es exitoso guarda en catalogo, crea la estructura
                GuardarPOMEstructura(sNombre);
               
            }
            catch (Exception ex)
            {
                
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }
        }


        public DataSet ObtenerIdPOM(string sNombre)
        {
            string query = "Select Id from CatPOM where Nombre ='"+ sNombre+"'";

            OleDbParameter[] opParametros = { };

            return fSQL(query, opParametros);
        }

        public DataSet ObtenerAñosPOM(decimal Id)
        {
            string query = "Select Nombre,Fecha_Inicio, Fecha_Fin from CatPOM where Id =" + Id;

            OleDbParameter[] opParametros = { };

            return fSQL(query, opParametros);
        }

        public DataSet ObtenerIdPOMPadre(decimal idPOM)
        {
            string query = "SELECT id AS idPadre FROM CatProceso where IdCatPom="+idPOM+" and Categoria='POM'";

            OleDbParameter[] opParametros = {  };

            return fSQL(query, opParametros);
        }


        public void GuardarPOMEstructura(string sNombre)
        {
            try
            {
                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";
                OleDbCommand comando;
                string q = "";
                con.Open();

                //OBTENER ID DE POM CREADO
                DataTable dt = ObtenerIdPOM(sNombre).Tables[0];
                decimal idPOM = 0;

                foreach (DataRow row in dt.Rows)
                {
                    idPOM = Convert.ToDecimal(row["Id"]);
                }

                //POM
                q = "INSERT INTO CatProceso (Nombre,idCatPOM_Ejecutores,Categoria,idCatPom,Padre) VALUES (@Nombre,@idCatPOM_Ejecutores,'POM',@idCatPom,0)";
                comando = new OleDbCommand(q, con);
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@Nombre", sNombre);
                comando.Parameters.AddWithValue("@idCatPOM_Ejecutores", idPOM);
                comando.Parameters.AddWithValue("@idCatPom", idPOM);
                comando.ExecuteNonQuery();
                con.Close();
                con.Open();
                //BUSCAR ID DEL PADRE POM
                DataTable dtPOM = ObtenerIdPOMPadre(idPOM).Tables[0];
                decimal idPadrePOM = 0;

                foreach (DataRow row in dtPOM.Rows)
                {
                    idPadrePOM = Convert.ToDecimal(row["idPadre"]);
                }
                //Categoria Ejecutores
                q = "INSERT INTO CatProceso (Nombre,idCatPOM_Ejecutores,Categoria,idCatPom,Padre) VALUES ('Ejecutores',"+ idPOM+",'Categoria',"+ idPOM+","+ idPadrePOM+")";
                comando = new OleDbCommand(q, con);
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@Padre", idPOM);
                comando.Parameters.AddWithValue("@idCatPOM_Ejecutores", idPOM);
                comando.Parameters.AddWithValue("@idCatPom", idPadrePOM);
                comando.ExecuteNonQuery();

                //Categoria Requerido
                q = "INSERT INTO CatProceso (Nombre,idCatPOM_Ejecutores,Categoria,idCatPom,Padre) VALUES ('Requerido'," + idPOM + ",'Categoria'," + idPOM + "," + idPadrePOM + ")";
                comando = new OleDbCommand(q, con);
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@Padre", idPOM);
                comando.Parameters.AddWithValue("@idCatPOM_Ejecutores", idPadrePOM);
                comando.Parameters.AddWithValue("@idCatPom", idPOM);
                comando.ExecuteNonQuery();

                //Categoria Peimer Ajuste
                q = "INSERT INTO CatProceso (Nombre,idCatPOM_Ejecutores,Categoria,idCatPom,Padre) VALUES ('Primer Ajuste'," + idPOM + ",'Categoria'," + idPOM + "," + idPadrePOM + ")";
                comando = new OleDbCommand(q, con);
                comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@Padre", idPOM);
                comando.Parameters.AddWithValue("@idCatPOM_Ejecutores", idPadrePOM);
                comando.Parameters.AddWithValue("@idCatPom", idPOM);
                comando.ExecuteNonQuery();

                con.Close();
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "Nuevo POM " + sNombre + " creado exitosamente";

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);

              

            }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }
        }

        public void GuardarCatActividades(RadGridView grid,string catalogo)
        {
            try
            {

                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();

                if(catalogo == "Actividades")
                {
                    string q = "INSERT INTO CatActividades (PRO_PRE,POZO,ACTIVIDAD,ACTIVIDAD_ESPECIFICA,No_EQUIPO,NOMBRE_DE_EQUIPO,ESTRUC,RENTA,OBJETIVO,INTERVALO,PROF,Fec_Inic,Fec_Term,Días,NOTAS,Campo16,LLAVE,LLAVE1,LLAVE2,LLAVE3,Campo21,ELEM_PEP,LLAVE_DE_CRUCE) VALUES (@pre,@pozo,@ACTIVIDAD,@ACTIVIDAD_ESPECIFICA,@No_EQUIPO,@NOMBRE_DE_EQUIPO,@ESTRUC,@RENTA,@OBJETIVO,@INTERVALO,@PROF,@Fec_Inic,@Fec_Term,@Días,@NOTAS,@Campo16,@LLAVE,@LLAVE1,@LLAVE2,@LLAVE3,@Campo21,@ELEM_PEP,@LLAVE_DE_CRUCE)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {

                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();

                        DataRowView row = (DataRowView)item;
                        DateTime fecha_i = Convert.ToDateTime(row[11].ToString());
                        DateTime fecha_f = Convert.ToDateTime(row[12].ToString());

                        //comando.Parameters.AddWithValue("@ID", Convert.ToDecimal(((GridViewCell)(row.Cells[0])).Value));
                        comando.Parameters.AddWithValue("@pre", row[0].ToString());
                        comando.Parameters.AddWithValue("@pozo", row[1].ToString());
                        comando.Parameters.AddWithValue("@ACTIVIDAD", row[2].ToString());
                        comando.Parameters.AddWithValue("@ACTIVIDAD_ESPECIFICA", row[3].ToString());
                        comando.Parameters.AddWithValue("@No_EQUIPO", row[4].ToString());
                        comando.Parameters.AddWithValue("@NOMBRE_DE_EQUIPO", row[5].ToString());
                        comando.Parameters.AddWithValue("@ESTRUC", row[6].ToString());
                        comando.Parameters.AddWithValue("@RENTA", row[7].ToString());
                        comando.Parameters.AddWithValue("@OBJETIVO", row[8].ToString());
                        comando.Parameters.AddWithValue("@INTERVALO", row[9].ToString());
                        comando.Parameters.AddWithValue("@PROF", row[10].ToString());
                        comando.Parameters.AddWithValue("@Fec_Inic", fecha_i);
                        comando.Parameters.AddWithValue("@Fec_Term", fecha_f);
                        comando.Parameters.AddWithValue("@Días", row[13].ToString());
                        comando.Parameters.AddWithValue("@Campo16", row[14].ToString());
                        comando.Parameters.AddWithValue("@NOTAS", row[15].ToString());
                        comando.Parameters.AddWithValue("@LLAVE", row[16].ToString());
                        comando.Parameters.AddWithValue("@LLAVE1", row[17].ToString());
                        comando.Parameters.AddWithValue("@LLAVE2", row[18].ToString());
                        comando.Parameters.AddWithValue("@LLAVE3", row[19].ToString());
                        comando.Parameters.AddWithValue("@Campo21", row[20].ToString());
                        comando.Parameters.AddWithValue("@ELEM_PEP", row[21].ToString());
                        comando.Parameters.AddWithValue("@LLAVE_DE_CRUCE", row[22].ToString());

                        comando.ExecuteNonQuery();
                    }
                }
                else if(catalogo== "Admon")
                {
                    string q = "INSERT INTO CatAdmon (Llave,Cve_Campo,ID_Asignacion,Des_Asignacion,Cve_Asignacion,Des_Campo,Cve_Proyecto,Des_Proyecto,Campo_Cab_Asig,Activo,Proy_Prog) VALUES (@Llave,@Cve_Campo,@ID_Asignacion,@Des_Asignacion,@Cve_Asignacion,@Des_Campo,@Cve_Proyecto,@Des_Proyecto,@Campo_Cab_Asig,@Activo,@Proy_Prog)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {


                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();

                        DataRowView row = (DataRowView)item;
                        

                        // comando.Parameters.AddWithValue("@Id", Convert.ToDecimal(((GridViewCell)(row.Cells[0])).Value));
                        comando.Parameters.AddWithValue("@Llave", row[0].ToString());
                        comando.Parameters.AddWithValue("@Cve_Campo", row[1].ToString());
                        comando.Parameters.AddWithValue("@ID_Asignacion", row[2].ToString());
                        comando.Parameters.AddWithValue("@Des_Asignacion", row[3].ToString());
                        comando.Parameters.AddWithValue("@Cve_Asignacion", row[4].ToString());
                        comando.Parameters.AddWithValue("@Des_Campo", row[5].ToString());
                        comando.Parameters.AddWithValue("@Cve_Proyecto", row[6].ToString());
                        comando.Parameters.AddWithValue("@Des_Proyecto", row[7].ToString());
                        comando.Parameters.AddWithValue("@Campo_Cab_Asig", row[8].ToString());
                        comando.Parameters.AddWithValue("@Activo", row[9].ToString());
                        comando.Parameters.AddWithValue("@Proy_Prog", row[10].ToString());

                        comando.ExecuteNonQuery();
                    }
                }
                else if (catalogo == "AsigCampoPYIN")
                {
                    string q = "INSERT INTO CatAsigCampoPYIN (Campo_Cab_Asig,Subdirección,ID_Asignacion,Cve_Asignacion,Des_Asignacion,Cve_Campo,Des_Campo,Cve_Proyecto,Des_Proyecto,Campo10,Activo,Estatus,F_inicio,F_fin) VALUES (@Campo_Cab_Asig,@Subdirección,@ID_Asignacion,@Cve_Asignacion,@Des_Asignacion,@Cve_Campo,@Des_Campo,@Cve_Proyecto,@Des_Proyecto,@Campo10,@Activo,@Estatus,@F_inicio,@F_fin)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {

                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();
                        DataRowView row = (DataRowView)item;

                        DateTime fecha_i = Convert.ToDateTime(row[12].ToString());
                        DateTime fecha_f = Convert.ToDateTime(row[13].ToString());


                        comando.Parameters.AddWithValue("@Campo_Cab_Asig", row[0].ToString());
                        comando.Parameters.AddWithValue("@Subdirección", row[1].ToString());
                        comando.Parameters.AddWithValue("@ID_Asignacion", row[2].ToString());
                        comando.Parameters.AddWithValue("@Cve_Asignacion", row[3].ToString());
                        comando.Parameters.AddWithValue("@Des_Asignacion", row[4].ToString());
                        //comando.Parameters.AddWithValue("@Subdirección", row[5].ToString());
                        comando.Parameters.AddWithValue("@Cve_Campo", row[5].ToString());
                        comando.Parameters.AddWithValue("@Des_Campo", row[6].ToString());
                        comando.Parameters.AddWithValue("@Cve_Proyecto", row[7].ToString());
                        comando.Parameters.AddWithValue("@Des_Proyecto", row[8].ToString());
                        comando.Parameters.AddWithValue("@Campo10", row[9].ToString());
                        comando.Parameters.AddWithValue("@Activo", row[10].ToString());
                        comando.Parameters.AddWithValue("@Estatus", row[11].ToString());
                        comando.Parameters.AddWithValue("@F_inicio", fecha_i);
                        comando.Parameters.AddWithValue("@F_fin", fecha_f);

                        comando.ExecuteNonQuery();
                    }
                }
                else if (catalogo == "Contratos")
                {
                    string q = "INSERT INTO CatContratos (CONTRATO,RESERVA,Descripción_Contrato,FecIncContrato,FecVencContrato,POLITICAPAGO,COMPAÑÍA,PROVEEDOR,CONCATENAR) VALUES (@CONTRATO,@RESERVA,@Descripción_Contrato,@FecIncCont,@FecVencContrato,@POLITICAPAGO,@COMPAÑÍA,@PROVEEDOR,@CONCATENAR)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    ///Sergio Lira 18/08/2017
                    ///Modificado para que guarde todos las filas del grid y no solo por paginado
                    var itemsSource = grid.ItemsSource as IEnumerable;
                    
                    foreach (var item in itemsSource)
                    {
                        
                        comando.Parameters.Clear();
                        
                        DataRowView row = (DataRowView)item;
                      
                        //Agregado porque causa conflicto si el campo esta vacion porque es numerico en la base de datos
                        var pol = row[5].ToString() == "" ? "0" : row[5].ToString();
                        var pro = row[7].ToString() == "" ? "0" : row[7].ToString();

                        comando.Parameters.AddWithValue("@CONTRATO", row[0].ToString());
                        comando.Parameters.AddWithValue("@RESERVA",  row[1].ToString());
                        comando.Parameters.AddWithValue("@Descripción_Contrato", row[2].ToString());
                        comando.Parameters.AddWithValue("@FecIncCont", Convert.ToDateTime(row[3].ToString()));
                        comando.Parameters.AddWithValue("@FecVencContrato", Convert.ToDateTime(row[4].ToString()));
                        comando.Parameters.AddWithValue("@POLITICAPAGO", pol);
                        comando.Parameters.AddWithValue("@COMPAÑÍA", row[6].ToString());
                        comando.Parameters.AddWithValue("@PROVEEDOR", pro);
                        comando.Parameters.AddWithValue("@CONCATENAR", row[8].ToString());

                        comando.ExecuteNonQuery();
                    }
                }
                else if (catalogo == "IdPozo")
                {
                    string q = "INSERT INTO CatIdPozo (Id,Pozo,IIP) VALUES (@Id,@Pozo,@IIP)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {

                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();
                        DataRowView row = (DataRowView)item;

                        Decimal id = row[0].ToString() == "" ? 0 : Convert.ToDecimal(row[0].ToString());

                        comando.Parameters.AddWithValue("@Id", id);
                        comando.Parameters.AddWithValue("@Pozo", row[1].ToString());
                        comando.Parameters.AddWithValue("@IIP", row[2].ToString());

                        comando.ExecuteNonQuery();
                    }
                }
                else if (catalogo == "LlaveControl")
                {
                    string q = "INSERT INTO CatLlaveControl (Id1,prott,PP,id_presentacion,ID,ID2) VALUES (@Id1,@prott,@PP,@id_presentacion,@ID,@ID2)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {

                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();
                        DataRowView row = (DataRowView)item;

                        Decimal id1 = row[0].ToString() == "" ? 0 : Convert.ToDecimal(row[0].ToString());

                        comando.Parameters.AddWithValue("@Id1", id1);
                        comando.Parameters.AddWithValue("@prott", row[1].ToString());
                        comando.Parameters.AddWithValue("@PP", row[2].ToString());
                        comando.Parameters.AddWithValue("@id_presentacion", row[3].ToString());
                        comando.Parameters.AddWithValue("@ID", row[4].ToString());
                        comando.Parameters.AddWithValue("@ID2", row[5].ToString());
                        
                        comando.ExecuteNonQuery();
                    }
                }
                else if (catalogo == "PPEquivalentes")
                {
                    string q = "INSERT INTO CatPPEquivalentes (Programa_anterior,Elemento_PEP_Anterior,Descripción_PP_Anterior,Programa_Vigente,Elemento_PEP,Descripción_PP) VALUES (@Programa_anterior,@Elemento_PEP_Anterior,@Descripción_PP_Anterior,@Programa_Vigente,@Elemento_PEP,@Descripción_PP)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {

                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();
                        DataRowView row = (DataRowView)item;

                        //comando.Parameters.AddWithValue("@Id", Convert.ToDecimal(((GridViewCell)(row.Cells[0])).Value));
                        comando.Parameters.AddWithValue("@Programa_anterior", row[0].ToString());
                        comando.Parameters.AddWithValue("@Elemento_PEP_Anterior", row[1].ToString());
                        comando.Parameters.AddWithValue("@Descripción_PP_Anterior", row[2].ToString());
                        comando.Parameters.AddWithValue("@Programa_Vigente", row[3].ToString());
                        comando.Parameters.AddWithValue("@Elemento_PEP", row[4].ToString());
                        comando.Parameters.AddWithValue("@Descripción_PP", row[5].ToString());
                        comando.ExecuteNonQuery();
                    }
                }
                else if (catalogo == "Proyectos")
                {
                    string q = "INSERT INTO CatProyectos (PROYECTO,PYIN) VALUES (@PROYECTO,@PYIN)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {

                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();
                        DataRowView row = (DataRowView)item;

                        comando.Parameters.AddWithValue("@PROYECTO", row[0].ToString());
                        comando.Parameters.AddWithValue("@PYIN", row[1].ToString());
                        comando.ExecuteNonQuery();
                    }
                }
                else if (catalogo == "POZOADMON")
                {
                    //string q = " INSERT INTO CatPOZOADMON (PROGRAMA,ELEMENTO_PEP, PRO_POZO, CONCATENAR,POZO,Soc_CO,Entidad_CP,Proy_Aval,PROY,EPEP_Descripcion,EPEP_Ind,EPEP_Cent, Prog_Inic,Prog_Term,Fecha,Prog,To,Campo,PTO) VALUES ( '41AA0100KG4Z1601', 'E41AA01Z16KG401', '41AKG4T. APOYO', '41AKG4', 'T. APOYO', 'PEP', 'PEP', 'PROD', '41A', 'B08 GTOS OPERAC EQS FUERA DE OPERN BALAM', '0', '2531', '01/10/2016', '31/12/2025', '', 'KG', '4', 'Z16', 'KG4')";

                    string q = " INSERT INTO CatPOZOADMON (PROGRAMA, ELEMENTO_PEP, PRO_POZO, CONCATENAR, POZO, Soc_CO, Entidad_CP, Proy_Aval, PROY, EPEP_Descripcion, EPEP_Ind, EPEP_Cent, Prog_Inic, Prog_Term, Fecha, Prog, Tox, Campo,PTO) VALUES(@PROGRAMA, @ELEMENTO_PEP, @PRO_POZO, @CONCATENAR, @POZO, @Soc_CO, @Entidad_CP, @Proy_Aval, @PROY, @EPEP_Descripcion, @EPEP_Ind, @EPEP_Cent, @Prog_Inic, @Prog_Term, @Fecha, @Prog, @Tox, @Campo, @PTO)";

                    OleDbCommand comando = new OleDbCommand(q, con);

                    //VALUES( '41AA0100KG4Z1601', 'E41AA01Z16KG401', '41AKG4T. APOYO', '41AKG4', 'T. APOYO', 'PEP', 'PEP', 'PROD', '41A', 'B08 GTOS OPERAC EQS FUERA DE OPERN BALAM',  '', 25310000, '01/10/2016', '31/12/2025', '', 'KG', '4', 'Z16', 'KG4'    )

                    /*VALUES(@PROGRAMA, @ELEMENTO_PEP, @PRO_POZO, @CONCATENAR, @POZO, @Soc_CO, @Entidad_CP, @Proy_Aval, @PROY, @EPEP_Descripcion,
                                @EPEP_Ind, @EPEP_Cent, @Prog_Inic, @Prog_Term, @Fecha, @Prog, @To, @Campo, @PTO)*/


                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {

                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();
                        DataRowView row = (DataRowView)item;

                        Decimal cent = row[11].ToString() == "" ? 0 : Convert.ToDecimal(row[11].ToString());
                        DateTime fi = Convert.ToDateTime(row[12].ToString()); // Convert.ToDateTime(row[11].ToString());
                        DateTime ff = Convert.ToDateTime(row[13].ToString()); // Convert.ToDateTime(row[12].ToString());


                        comando.Parameters.AddWithValue("@PROGRAMA", row[0].ToString());
                        comando.Parameters.AddWithValue("@ELEMENTO_PEP", row[1].ToString());
                        comando.Parameters.AddWithValue("@PRO_POZO", row[2].ToString());
                        comando.Parameters.AddWithValue("@CONCATENAR", row[3].ToString());
                        comando.Parameters.AddWithValue("@POZO", row[4].ToString());
                        comando.Parameters.AddWithValue("@Soc_CO", row[5].ToString());
                        comando.Parameters.AddWithValue("@Entidad_CP", row[6].ToString());
                        comando.Parameters.AddWithValue("@Proy_Aval", row[7].ToString());
                        comando.Parameters.AddWithValue("@PROY", row[8].ToString());
                        comando.Parameters.AddWithValue("@EPEP_Descripcion", row[9].ToString());
                        comando.Parameters.AddWithValue("@EPEP_Ind", row[10].ToString());
                        comando.Parameters.AddWithValue("@EPEP_Cent", cent);
                        comando.Parameters.AddWithValue("@Prog_Inic", fi);
                        comando.Parameters.AddWithValue("@Prog_Term", ff);
                        comando.Parameters.AddWithValue("@Fecha", row[14].ToString());
                        comando.Parameters.AddWithValue("@Prog", row[15].ToString());
                        comando.Parameters.AddWithValue("@Tox", row[16].ToString());
                        comando.Parameters.AddWithValue("@Campo", row[17].ToString());
                        comando.Parameters.AddWithValue("@PTO", row[18].ToString());

                        comando.ExecuteNonQuery();
                    }
                }
                else if (catalogo == "Supervisores")
                {
                    string q = "INSERT INTO CatSupervisores (USUARIO,FICHA) VALUES (@USUARIO,@FICHA)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    var itemsSource = grid.ItemsSource as IEnumerable;

                    foreach (var item in itemsSource)
                    {

                        //var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();
                        DataRowView row = (DataRowView)item;

                        comando.Parameters.AddWithValue("@USUARIO", row[0].ToString());
                        comando.Parameters.AddWithValue("@FICHA", row[1].ToString());
                        comando.ExecuteNonQuery();
                    }
                }

                con.Close();
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "El Cat"+catalogo+ " fué actualizado con exito.";

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);

            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }
        #endregion

        #region EJECUTORES
        public DataSet ObtenerProceso(Decimal idPOM)
        {
            string query = "Select id, Padre,Nombre,idCatPOM_Ejecutores,Categoria from CatProceso where idCatPom =" + idPOM;

            OleDbParameter[] opParametros = { };

            return fSQL(query, opParametros);
        }

        public DataSet ObtenerIdEjecutor(string sNombre, decimal idPOM)
        {
            string query = "SELECT id FROM CatEjecutores where Nombre = @Nombre and IdPOM=@idPOM";

            OleDbParameter[] opParametros = { new OleDbParameter("@Nombre", sNombre), new OleDbParameter("@idPOM", idPOM) };

            return fSQL(query, opParametros);
        }


        public DataSet ObtenerNombreEjecutor(decimal id)
        {
            string query = "SELECT Nombre FROM CatEjecutores where id=@id";

            OleDbParameter[] opParametros = { new OleDbParameter("@id", id) };

            return fSQL(query, opParametros);
        }

        public DataSet ObtenerIdEjecutorPadre(decimal idPOM)
        {
            string query = "SELECT id AS idPadre FROM CatProceso where IdCatPom=@idPOM and Nombre='Ejecutores' and Categoria='Categoria'";

            OleDbParameter[] opParametros = { new OleDbParameter("@idPOM", idPOM) };

            return fSQL(query, opParametros);
        }

        public void GuardaCatEjecutores(decimal idPOM, string nombreEjecutor)
        {
                        string q = "";
        OleDbConnection con = new OleDbConnection();
        OleDbCommand comando;
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();


                //INSERTA EL EJECUTOR EN EL CATALOGO
                q = "INSERT INTO CatEjecutores (Nombre,IdPOM) VALUES (@Nombre,@IdPOM)";
                comando = new OleDbCommand(q, con);
        comando.Parameters.Clear();
                comando.Parameters.AddWithValue("@Nombre", nombreEjecutor);
                comando.Parameters.AddWithValue("@IdPOM", idPOM);

                comando.ExecuteNonQuery();
            con.Close();
        }


        public void GuardaCatProceso(decimal idPOM, string nombreEjecutor, decimal idEjecutores)
        {
            string q = "";
            OleDbConnection con = new OleDbConnection();
            OleDbCommand comando;
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

            con.Open();

            //Busca el padre en el arbol para ejecutores
            DataTable dtIdEjecutoresPAdre = ObtenerIdEjecutorPadre(idPOM).Tables[0];
            decimal idEjecutoresPadre = 0;

            foreach (DataRow row in dtIdEjecutoresPAdre.Rows)
            {
                idEjecutoresPadre = Convert.ToDecimal(row["IdPadre"]);
            }


            //INSERTA NOMBRE DE EJECUTOR EN EL ARBOL DEL PROCESO
            q = "INSERT INTO CatProceso (Nombre,IdCatPom,Padre,Categoria,idCatPOM_Ejecutores) VALUES (@Nombre,@IdPOM,@Padre,'Hijo',@idCatPOM_Ejecutores)";
            comando = new OleDbCommand(q, con);
            comando.Parameters.Clear();
            comando.Parameters.AddWithValue("@Nombre", nombreEjecutor);
            comando.Parameters.AddWithValue("@IdPOM", idPOM);
            comando.Parameters.AddWithValue("@Padre", idEjecutoresPadre);
            comando.Parameters.AddWithValue("@idCatPOM_Ejecutores", idEjecutores);
            comando.ExecuteNonQuery();


            con.Close();

        }




        public void GuardarEjecutor(RadGridView grid, decimal idPOM,string nombreEjecutor)
        {
            try
            {

                GuardaCatEjecutores(idPOM, nombreEjecutor);
                

                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();


                //BUSCA EL ID ASIGNADO EN EL CATALGO DE EJECUTORES
                DataTable dt = ObtenerIdEjecutor(nombreEjecutor,idPOM).Tables[0];
                decimal idEjecutores=0;

                foreach (DataRow row in dt.Rows)
                    {
                    idEjecutores = Convert.ToDecimal(row["id"]);
                    }


                GuardaCatProceso(idPOM, nombreEjecutor,idEjecutores);

                //INSERTA DATOS DEL EXCEL DEL EJECUTOR
                q = "INSERT INTO Ejecutores (Id_Ejecutor,Id_POM,NO_CONTRATO,NO_RESERVA,FICHA_SUPERVISOR,AÑO,MOV_EQ,EQUIPO,PLATAFORMA,POZO,ACTIVIDAD,ACT_ESPECIFICA,FECHA_I,FECHA_T,C_GESTOR,PROY,PROGRAMA_PRESUPUESTARIO,POSPRE,MONEDA,ID_REGISTRO,ENE,FEB,MAR,ABR,MAY,JUN,JUL,AGO,SEP,OCT,NOV,DIC,TOTAL) VALUES (@Id_Ejecutor,@Id_POM,@NO_CONTRATO,@NO_RESERVA,@FICHA_SUPERVISOR,@AÑO,@MOV_EQ,@EQUIPO,@PLATAFORMA,@POZO,@ACTIVIDAD,@ACT_ESPECIFICA,@FECHA_I,@FECHA_T,@C_GESTOR,@PROY,@PROGRAMA_PRESUPUESTARIO,@POSPRE,@MONEDA,@ID_REGISTRO,@ENE,@FEB,@MAR,@ABR,@MAY,@JUN,@JUL,@AGO,@SEP,@OCT,@NOV,@DIC,@TOTAL)";
                    comando = new OleDbCommand(q, con);

                    foreach (var item in grid.Items)
                    {

                        var row = grid.ItemContainerGenerator.ContainerFromItem(item) as GridViewRow;
                        comando.Parameters.Clear();
                        comando.Parameters.AddWithValue("@Id_Ejecutor", idEjecutores);
                        comando.Parameters.AddWithValue("@Id_POM", idPOM);
                        comando.Parameters.AddWithValue("@NO_CONTRATO", Convert.ToString(((GridViewCell)(row.Cells[0])).Value));
                        comando.Parameters.AddWithValue("@NO_RESERVA", Convert.ToString(((GridViewCell)(row.Cells[1])).Value));
                        comando.Parameters.AddWithValue("@FICHA_SUPERVISOR", Convert.ToString(((GridViewCell)(row.Cells[2])).Value));
                        comando.Parameters.AddWithValue("@AÑO", Convert.ToString(((GridViewCell)(row.Cells[3])).Value));
                        comando.Parameters.AddWithValue("@MOV_EQ", Convert.ToString(((GridViewCell)(row.Cells[4])).Value));
                        comando.Parameters.AddWithValue("@EQUIPO", Convert.ToString(((GridViewCell)(row.Cells[5])).Value));
                        comando.Parameters.AddWithValue("@PLATAFORMA", Convert.ToString(((GridViewCell)(row.Cells[6])).Value));
                        comando.Parameters.AddWithValue("@POZO", Convert.ToString(((GridViewCell)(row.Cells[7])).Value));
                        comando.Parameters.AddWithValue("@ACTIVIDAD", Convert.ToString(((GridViewCell)(row.Cells[8])).Value));
                        comando.Parameters.AddWithValue("@ACT_ESPECIFICA", Convert.ToString(((GridViewCell)(row.Cells[9])).Value));

                    comando.Parameters.AddWithValue("@FECHA_I", Convert.ToDateTime(((GridViewCell)(row.Cells[10])).Value));
                    comando.Parameters.AddWithValue("@FECHA_T", Convert.ToDateTime(((GridViewCell)(row.Cells[11])).Value));
                    comando.Parameters.AddWithValue("@C_GESTOR", Convert.ToString(((GridViewCell)(row.Cells[12])).Value));
                    comando.Parameters.AddWithValue("@PROY", Convert.ToString(((GridViewCell)(row.Cells[13])).Value));
                    comando.Parameters.AddWithValue("@PROGRAMA_PRESUPUESTARIO", Convert.ToString(((GridViewCell)(row.Cells[14])).Value));
                    comando.Parameters.AddWithValue("@POSPRE", Convert.ToString(((GridViewCell)(row.Cells[15])).Value));
                    comando.Parameters.AddWithValue("@MONEDA", Convert.ToString(((GridViewCell)(row.Cells[16])).Value));
                    comando.Parameters.AddWithValue("@ID_REGISTRO", Convert.ToString(((GridViewCell)(row.Cells[17])).Value));
                    comando.Parameters.AddWithValue("@ENE", Convert.ToDecimal(((GridViewCell)(row.Cells[18])).Value));
                    comando.Parameters.AddWithValue("@FEB", Convert.ToDecimal(((GridViewCell)(row.Cells[19])).Value));
                    comando.Parameters.AddWithValue("@MAR", Convert.ToDecimal(((GridViewCell)(row.Cells[20])).Value));
                    comando.Parameters.AddWithValue("@ABR", Convert.ToDecimal(((GridViewCell)(row.Cells[21])).Value));
                    comando.Parameters.AddWithValue("@MAY", Convert.ToDecimal(((GridViewCell)(row.Cells[22])).Value));
                    comando.Parameters.AddWithValue("@JUN", Convert.ToDecimal(((GridViewCell)(row.Cells[23])).Value));
                    comando.Parameters.AddWithValue("@JUL", Convert.ToDecimal(((GridViewCell)(row.Cells[24])).Value));
                    comando.Parameters.AddWithValue("@AGO", Convert.ToDecimal(((GridViewCell)(row.Cells[25])).Value));
                    comando.Parameters.AddWithValue("@SEP", Convert.ToDecimal(((GridViewCell)(row.Cells[26])).Value));
                    comando.Parameters.AddWithValue("@OCT", Convert.ToDecimal(((GridViewCell)(row.Cells[27])).Value));
                    comando.Parameters.AddWithValue("@NOV", Convert.ToDecimal(((GridViewCell)(row.Cells[28])).Value));
                    comando.Parameters.AddWithValue("@DIC", Convert.ToDecimal(((GridViewCell)(row.Cells[29])).Value));
                    comando.Parameters.AddWithValue("@TOTAL", Convert.ToDecimal(((GridViewCell)(row.Cells[30])).Value));



                    comando.ExecuteNonQuery();
                    }

            

                con.Close();
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "El Ejecutor fué guardado con exito.";

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);

            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public DataSet ObtenerEjecutor(decimal idEjecutor)
        {
            string query = "SELECT NO_CONTRATO,NO_RESERVA,FICHA_SUPERVISOR,AÑO,MOV_EQ,EQUIPO,PLATAFORMA,POZO,ACTIVIDAD,ACT_ESPECIFICA,FECHA_I,FECHA_T,C_GESTOR,PROY,PROGRAMA_PRESUPUESTARIO,POSPRE,MONEDA,ID_REGISTRO,ENE,FEB,MAR,ABR,MAY,JUN,JUL,AGO,SEP,OCT,NOV,DIC,TOTAL FROM Ejecutores WHERE Id_Ejecutor=@idEjecutor";

            OleDbParameter[] opParametros = {new OleDbParameter("@idEjecutor", idEjecutor)};

            return fSQL(query, opParametros);
        }


      


        public DataSet VerificaEjecutorCalendarizado(decimal idEjecutor, string nombreTabla)
        {
            string query = "SELECT * FROM "+nombreTabla+ " WHERE idEjecutores=@idEjecutor;";

            OleDbParameter[] opParametros = { new OleDbParameter("@idEjecutor", idEjecutor) };

            return fSQL(query, opParametros);
        }



        public void CalendarizarEjecutor(decimal idEjecutor,decimal idPOM,decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {
                
               


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"INSERT INTO "+sNombre+" ( Base, Contrato, Reserva, MovEquipos, StatusOriginal, Ficha, Suptcia, Activo, Pyin, ClasifGral, Prog, ClasifAct, ClasifDeta, Ejecutores, Fondo, CT, Depto, CTEjec, DeptoEjec, NuevoCGestor, ProPre, RG, CO, Pozo, ACTIVIDADESPECIFICA, PospreOrig, Moneda, DENE"+FechaIni+ ", DFEB" + FechaIni + ", DMAR" + FechaIni + ", DABR" + FechaIni + ", DMAY" + FechaIni + ", DJUN" + FechaIni + ", DJUL" + FechaIni + ", DAGO" + FechaIni + ", DSEP" + FechaIni + ", DOCT" + FechaIni + ", DNOV" + FechaIni + ", DDIC" + FechaIni + ", TOTDev" + FechaIni + ", FEne" + FechaIni + ", FFeb" + FechaIni + ", FMar" + FechaIni + ", FAbr" + FechaIni + ", FMay" + FechaIni + ", FJun" + FechaIni + ", FJul" + FechaIni + ", FAgo" + FechaIni + ", FSep" + FechaIni + ", FOct" + FechaIni + ", FNov" + FechaIni + ", FDic" + FechaIni + ", TOTFe" + FechaIni + ", PIVDEV, FEC_INIC, FEC_TERM, Días,idEjecutores,idPOM ) SELECT 'BD Req' AS Base, Ejecutores.[NO_CONTRATO], Ejecutores.[NO_RESERVA], Ejecutores.[MOV_EQ], 'ProyEjecutor', Ejecutores.[FICHA_SUPERVISOR], 'ECPM' AS Suptcia, '253 Activo Cantarell' AS Activo, LEFT(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 3) AS Pyin, SWITCH(MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'QA', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'PH', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'PD', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 3) = 'KGZ', 'MO', True, 'ADMON') AS ClasifGral, MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 3) AS Prog, SWITCH(MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'QA', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'PH', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'PD', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 3) = 'KGZ', 'MO', True, 'ADMON') AS ClasifAct, SWITCH(MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'QA', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'PH', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 2) = 'PD', 'POZOS', MID(Ejecutores.[PROGRAMA_PRESUPUESTARIO], 9, 3) = 'KGZ', 'MO', True, 'ADMON') AS ClasifDeta, SWITCH(Ejecutores.[C_Gestor] = '20928100', 'SML GIC', Ejecutores.[C_Gestor] = '20929000', 'SML GLM', Ejecutores.[C_Gestor] = '20928000', 'SML GMI', Ejecutores.[C_Gestor] = '22156800', 'SSE', Ejecutores.[C_Gestor] = '22121000', 'GCO', LEFT(Ejecutores.[C_Gestor], 3) = '223', 'PPS', True, Ejecutores.[C_Gestor]) AS Ejecutores, 'PEF' AS Fondo, LEFT(Ejecutores.[C_GESTOR], 3) AS CT, RIGHT(Ejecutores.[C_GESTOR], 5) AS Depto, SWITCH(Ejecutores.[C_GESTOR] = '20928100', 'SML GIC', Ejecutores.[C_Gestor] = '20929000', 'SML GLM', Ejecutores.[C_Gestor] = '20928000', 'SML GMI', Ejecutores.[C_Gestor] = '22156800', 'SSE', Ejecutores.[C_Gestor] = '22121000', 'GCO', LEFT(Ejecutores.[C_Gestor], 3) = '223', 'PPS', True, Ejecutores.[C_Gestor]) AS CTEjec, SWITCH(Ejecutores.[C_Gestor] = '22156800', 'SSE', LEFT(Ejecutores.[C_Gestor], 3) = '223', 'PPS', LEFT(Ejecutores.[C_Gestor], 3) = '209', 'SML', Ejecutores.[C_Gestor] = '22121000', 'GCO', True, Ejecutores.[C_Gestor]) AS DeptoEjec, Ejecutores.[C_GESTOR], Ejecutores.[PROGRAMA_PRESUPUESTARIO], LEFT(Ejecutores.POSPRE,3) AS RG, RIGHT(Ejecutores.POSPRE, 6) AS CO, Ejecutores.POZO AS Pozo, Ejecutores.[ACT_ESPECIFICA], Ejecutores.POSPRE, SWITCH(Ejecutores.MONEDA= 'USM','2', Ejecutores.MONEDA= 'MXP','1', Ejecutores.MONEDA= 'USD','2', True, Ejecutores.MONEDA) AS Moneda, SWITCH([Ejecutores.ENE] IS NULL, 0, TRUE, [Ejecutores.ENE]) AS Ene, SWITCH([Ejecutores.FEB] IS NULL, 0, TRUE, [Ejecutores.FEB]) AS Feb, SWITCH([Ejecutores.MAR] IS NULL, 0, TRUE, [Ejecutores.MAR]) AS Mar, SWITCH([Ejecutores.ABR] IS NULL, 0, TRUE, [Ejecutores.ABR]) AS Abr, SWITCH([Ejecutores.MAY] IS NULL, 0, TRUE, [Ejecutores.MAY]) AS May, SWITCH([Ejecutores.JUN] IS NULL, 0, TRUE, [Ejecutores.JUN]) AS Jun, SWITCH([Ejecutores.JUL] IS NULL, 0, TRUE, [Ejecutores.JUL]) AS Jul, SWITCH([Ejecutores.AGO] IS NULL, 0, TRUE, [Ejecutores.AGO]) AS Ago, SWITCH([Ejecutores.SEP] IS NULL, 0, TRUE, [Ejecutores.SEP]) AS Sep, SWITCH([Ejecutores.OCT] IS NULL, 0, TRUE, [Ejecutores.OCT]) AS Oct, SWITCH([Ejecutores.NOV] IS NULL, 0, TRUE, [Ejecutores.NOV]) AS Nov, SWITCH([Ejecutores.DIC] IS NULL, 0, TRUE, [Ejecutores.DIC]) AS Dic,(SWITCH([Ejecutores.ENE] IS NULL, 0, TRUE, [Ejecutores.ENE])+SWITCH([Ejecutores.FEB] IS NULL, 0, TRUE, [Ejecutores.FEB])+SWITCH([Ejecutores.MAR] IS NULL, 0, TRUE, [Ejecutores.MAR])+SWITCH([Ejecutores.ABR] IS NULL, 0, TRUE, [Ejecutores.ABR])+SWITCH([Ejecutores.MAY] IS NULL, 0, TRUE, [Ejecutores.MAY])+SWITCH([Ejecutores.JUN] IS NULL, 0, TRUE, [Ejecutores.JUN])+SWITCH([Ejecutores.JUL] IS NULL, 0, TRUE, [Ejecutores.JUL])+SWITCH([Ejecutores.AGO] IS NULL, 0, TRUE, [Ejecutores.AGO])+SWITCH([Ejecutores.SEP] IS NULL, 0, TRUE, [Ejecutores.SEP])+SWITCH([Ejecutores.OCT] IS NULL, 0, TRUE, [Ejecutores.OCT])+SWITCH([Ejecutores.NOV] IS NULL, 0, TRUE, [Ejecutores.NOV])+SWITCH([Ejecutores.DIC] IS NULL, 0, TRUE, [Ejecutores.DIC])) AS DevTotal, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Ejecutores.[ID_REGISTRO], Ejecutores.[FECHA_I], Ejecutores.[FECHA_T], Ejecutores.[FECHA_T] - Ejecutores.[FECHA_I] AS Dias," + idEjecutor + "," + idPOM + " FROM Ejecutores WHERE Id_Ejecutor=" + idEjecutor;
                comando = new OleDbCommand(q, con);

            
                  

                    comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS "+ ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor1(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {




                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+@" INNER JOIN CatPOZOADMON ON ("+sNombre+@".PROPRE=[CatPOZOADMON].PROGRAMA) AND ("+sNombre+@".POZO=[CatPOZOADMON].POZO) SET "+sNombre+@".ElementoPEP = [CatPOZOADMON].[ELEMENTO_PEP]
                    WHERE "+sNombre+@".ElementoPEP Is Null And "+sNombre+@".STATUSORIGINAL Like 'ProyEjecutor' AND "+sNombre+@".IdEjecutores = "+idEjecutor+" AND "+sNombre+@".IdPOM ="+idPOM;
                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA ELEMENTO PEP EJECUTORES CALENDARIZADOS 1 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor2(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {




                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+@" SET "+sNombre+@".ElementoPEP = 'ERROR'
                    WHERE (((["+sNombre+@"].[ElementoPEP]) Is Null))
                    AND "+sNombre+@".STATUSORIGINAL like 'ProyEjecutor' AND "+sNombre+@".IdEjecutores = " + idEjecutor + " AND "+sNombre+@".IdPOM =" + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA ELEMENTO PEP EJECUTORES CALENDARIZADOS 2 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor3(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {




                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+@" INNER JOIN CatPOZOADMON ON "+sNombre+@".PROPRE=[CatPOZOADMON].PROGRAMA SET "+sNombre+@".ElementoPEP = CatPOZOADMON.[ELEMENTO_PEP], "+sNombre+@".pozo = [CatPOZOADMON].pozo
                      WHERE ((("+sNombre+@".ElementoPEP) Is Null)) And "+sNombre+@".STATUSORIGINAL Like 'ProyEjecutor' AND "+sNombre+@".IdEjecutores = " + idEjecutor + " AND "+sNombre+@".IdPOM =" + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA ELEMENTO PEP EJECUTORES CALENDARIZADOS 3 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor4(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {




                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+@" INNER JOIN CatActividades ON ("+sNombre+@".PROPRE=CatActividades.[PRO_PRE]) AND (["+sNombre+@"].Pozo = [CatActividades].Pozo) SET "+sNombre+@".Actividad = CatActividades.Actividad, "+sNombre+ @".ActividadEspecifica = CatActividades.[ACTIVIDAD_ESPECIFICA], " + sNombre + @".NoEQUIPO = CatActividades.[No_EQUIPO], " + sNombre + @".NOMBREEQUIPO = CatActividades.[NOMBRE_DE_EQUIPO], " + sNombre + @".ESTRUC = CatActividades.ESTRUC, " + sNombre + @".RENTA = CatActividades.RENTA, " + sNombre + @".OBJETIVO = CatActividades.OBJETIVO, " + sNombre + @".INTERVALO = CatActividades.INTERVALO, " + sNombre + @".PROF = CatActividades.PROF, " + sNombre + @".Fec_Inic = CatActividades.Fec_Inic, " + sNombre + @".Fec_Term = CatActividades.Fec_Term, " + sNombre + @".Días = CatActividades.Días, " + sNombre + @".Notas = CatActividades.NOTAS
                    WHERE " + sNombre + @".STATUSORIGINAL Like 'ProyEjecutor' or " + sNombre + @".STATUSORIGINAL='Proyección Pozos' AND " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM =" + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA ACTIVIDAD EJECUTORES CALENDARIZADOS 4 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor5(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {




                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" INNER JOIN CatPOZOADMON ON [" + sNombre + @"].ElementoPEP=CatPOZOADMON.[ELEMENTO_PEP] SET " + sNombre + @".DescElemPep = CatPOZOADMON.[EPEP:_Descripción]
                    WHERE " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM =" + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA DESCRIPCION ELEMENTO PEP EJECUTORES CALENDARIZADOS 5 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }




        public void CalendarizarEjecutor6(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+@" INNER JOIN CatActividades ON "+sNombre+@".PROPRE=CatActividades.[PRO_PRE] SET "+sNombre+@".Actividad = CatActividades.Actividad, "+sNombre+@".ActividadEspecifica = CatActividades.[ACTIVIDAD_ESPECIFICA], "+sNombre+@".NoEQUIPO = CatActividades.[No_EQUIPO], "+sNombre+@".NOMBREEQUIPO = CatActividades.[NOMBRE_DE_EQUIPO], "+sNombre+@".ESTRUC = CatActividades.ESTRUC, "+sNombre+@".RENTA = CatActividades.RENTA, "+sNombre+@".OBJETIVO = CatActividades.OBJETIVO, "+sNombre+@".INTERVALO = CatActividades.INTERVALO, "+sNombre+ @".PROF = CatActividades.PROF,  " + sNombre + @".Fec_Inic = CatActividades.Fec_Inic,  " + sNombre + @".Fec_Term = CatActividades.Fec_Term,  " + sNombre + @".Días = CatActividades.Días,  " + sNombre + @".Notas = CatActividades.NOTAS
                    WHERE (((Mid([" + sNombre + @"].PROPRE,9,3)) Not In ('PH0','PH5','PH7','PH9'))) AND  " + sNombre + @".IdEjecutores = " + idEjecutor + " AND  " + sNombre + @".IdPOM =" + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA DESCRIPCION ELEMENTO PEP EJECUTORES CALENDARIZADOS 6 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor7(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" SET Actividad = SWITCH(MID(ProPre,9,3)='PH0','RMA',MID(ProPre,9,3)='PH5','RME',MID(ProPre,9,3)='PH7','EST',MID(ProPre,9,3)='PH9','TIN',True,'*')
                       WHERE Actividad is null AND " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM =" + idPOM; 

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA ACTIVIDAD PH EJECUTORES CALENDARIZADOS 7 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }

        public void CalendarizarEjecutor8(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" SET ClasifGral = SWITCH(LEFT(PROG,2)='QA','POZOS',LEFT(PROG,2)='PH','POZOS',LEFT(PROG,2)='PD','POZOS',LEFT(PROG,2)='PA','POZOS',LEFT(PROG,2)='PF','POZOS',LEFT(PROG,3)='KGZ','MO',True,'ADMON'), ClasifAct = SWITCH(LEFT(PROG,2)='QA','POZOS',LEFT(PROG,2)='PH','POZOS',LEFT(PROG,2)='PD','POZOS',LEFT(PROG,2)='PA','POZOS',LEFT(PROG,2)='PF','POZOS',LEFT(PROG,3)='KGZ','MO',True,'ADMON'), ClasifDeta = SWITCH(LEFT(PROG,2)='QA','POZOS',LEFT(PROG,2)='PH','POZOS',LEFT(PROG,2)='PD','POZOS',LEFT(PROG,2)='PA','POZOS',LEFT(PROG,2)='PF','POZOS',LEFT(PROG,3)='KGZ','MO',True,'ADMON'), Ejecutores = SWITCH(NuevoCgestor='20928100','SML GIC',NuevoCgestor='20929000','SML GLM',NuevoCgestor='20928000','SML GMI',NuevoCgestor='20919000','SML GAM',NuevoCgestor='22156800','SSE',LEFT(NuevoCgestor,3)='223','PPS',LEFT(NuevoCgestor,3)='232','UNP',True,NuevoCgestor), CTEjec = SWITCH(NuevoCgestor='20928100','SML GIC',NuevoCgestor='20929000','SML GLM',NuevoCgestor='20928000','SML GMI',NuevoCgestor='20919000','SML GAM',NuevoCgestor='22156800','SSE',LEFT(NuevoCgestor,3)='223','PPS',LEFT(NuevoCgestor,3)='232','UNP',True,NuevoCgestor), DeptoEjec = SWITCH(NuevoCgestor='22156800','SSE',LEFT(NuevoCgestor,3)='223','PPS',LEFT(NuevoCgestor,3)='232','UNP',LEFT(NuevoCgestor,3)='209','SML',True,NuevoCgestor)
                        WHERE " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM =" + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA CLASIF EJECUTOR EJECUTORES CALENDARIZADOS 8 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }

        public void CalendarizarEjecutor9(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" INNER JOIN CatProyectos ON [" + sNombre + @"].PYIN=CatProyectos.PYIN SET " + sNombre + @".Proyecto = CatProyectos.PROYECTO
                     WHERE " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA PROYECTOS EJECUTORES CALENDARIZADOS 9 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }



        public void CalendarizarEjecutor10(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" INNER JOIN CatLlaveControl ON [" + sNombre + @"].PROG=CatLlaveControl.[pro-tt] SET " + sNombre + @".Llavecontrol = [" + sNombre + @"].PYIN & 'ZZZZZ' & CatLlaveControl.PP & 'ZZZZZZ'
                        WHERE " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA LLLAVE CONTROL EJECUTORES CALENDARIZADOS 10 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }




        public void CalendarizarEjecutor11(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" SET CO = Right([" + sNombre + @"].POSPREORIG,6)
                    WHERE ((([" + sNombre + @"].CO) Is Null)) AND " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;


                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA CON ORIGEN EJECUTORES CALENDARIZADOS 12 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }

        public void CalendarizarEjecutor12(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" SET Cve_Campo = Mid(PROPRE,12,3)
                        WHERE statusoriginal <> 'PROY MO' AND " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;


                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA CLAVE CAMPO EJECUTORES CALENDARIZADOS 12 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor13(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" INNER JOIN CatPOZOADMON ON [" + sNombre + @"].PROPRE=CatPOZOADMON.PROGRAMA SET " + sNombre + @".pozo = CatPOZOADMON.pozo, " + sNombre + @".elementopep = CatPOZOADMON.[elemento_pep], fec_inic = '01/01/2015', Fec_Term = '31/12/2015', Días = 365, Actividad = '*', NoEQUIPO = '*', NOMBREEQUIPO = '*', ESTRUC = '*', RENTA = '*', OBJETIVO = '*', INTERVALO = '*', PROF = '*', Notas = '*'
                       WHERE [" + sNombre + @"].elementopep Is Null And statusoriginal Like 'ProyEjecutor*' AND " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;


                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA ELEMENTO PEP EJECUTORES CALENDARIZADOS 13 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }

        public void CalendarizarEjecutor14(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" SET contrato = mid(contrato,1,9)
                    WHERE len(contrato)=10 AND " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA CONTRATO EJECUTORES CALENDARIZADOS 14 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor15(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE CatContratos INNER JOIN "+sNombre+ @" ON CatContratos.CONTRATO = " + sNombre + @".CONTRATO SET " + sNombre + @".[Descripción_Contrato] = CatContratos.[Descripción_Contrato], " + sNombre + @".Reserva = CatContratos.[Reserva], " + sNombre + @".FecVencContrato = CatContratos.[FecVencContrato], " + sNombre + @".politicapago = CatContratos.PoliticaPago, " + sNombre + @".FecIniContrato = [CatContratos].[FecIncContrato], " + sNombre + @".Compañia = [CatContratos].[COMPAÑÍA], " + sNombre + @".Acreedor = [CatContratos].[PROVEEDOR]
                WHERE " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA DESCRIPCION CONTRATO EJECUTORES CALENDARIZADOS 15 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }




        public void CalendarizarEjecutor16(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" INNER JOIN CatAsigCampoPYIN ON [" + sNombre + @"].Cve_Campo=CatAsigCampoPYIN.Cve_Campo SET " + sNombre + @".Id_Asignacion = CatAsigCampoPYIN.id_asignacion, " + sNombre + @".DescripcionAsignacion = [CatAsigCampoPYIN].[Des_Asignacion], " + sNombre + @".Cve_Asignacion = CatAsigCampoPYIN.Cve_Asignacion
                    WHERE " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA ID ASIGNACION EJECUTORES CALENDARIZADOS 16 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }



        public void CalendarizarEjecutor17(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombre+ @" SET " + sNombre + @".PoliticaPago = '20'
                    WHERE ((([" + sNombre + @"].politicaPago) Is Null)) And statusoriginal Like 'ProyEjecutor*' AND " + sNombre + @".IdEjecutores = " + idEjecutor + " AND " + sNombre + @".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA POLITICA PAGO NULA EJECUTORES CALENDARIZADOS 17 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }

        public void CalendarizarEjecutor18(decimal idEjecutor, decimal idPOM, decimal FechaIni,decimal FechaFin,string nombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();


                
                q = @"UPDATE "+nombrePOM+@" SET 
                    TOTDev"+FechaIni+" = IIF(IsNull(DENE"+FechaIni+ "),0,DENE" + FechaIni + ")+IIF(IsNull(DFEB" + FechaIni+ "),0,DFEB" + FechaIni + ")+IIF(IsNull(DMAR" + FechaIni+ "),0,DMAR" + FechaIni + ")+IIF(IsNull(DABR" + FechaIni+ "),0,DABR" + FechaIni + ")+IIF(IsNull(DMAY" + FechaIni+ "),0,DMAY" + FechaIni + ")+IIF(IsNull(DJUN" + FechaIni+ "),0,DJUN" + FechaIni + ")+IIF(IsNull(DJUL" + FechaIni+ "),0,DJUL" + FechaIni + ")+IIF(IsNull(DAGO" + FechaIni+ "),0,DAGO" + FechaIni + ")+IIF(IsNull(DSEP" + FechaIni+ "),0,DSEP" + FechaIni + ")+IIF(IsNull(DOCT" + FechaIni+ "),0,DOCT" + FechaIni + ")+IIF(IsNull(DNOV" + FechaIni+ "),0,DNOV" + FechaIni + ")+IIF(IsNull(DDIC" + FechaIni+ "),0,DDIC" + FechaIni + @"), 
                    TOTFe" + FechaIni+ " = IIF(IsNull(FENE" + FechaIni + "),0,FENE" + FechaIni + ")+IIF(IsNull(FFEB" + FechaIni + "),0,FFEB" + FechaIni + ")+IIF(IsNull(FMAR" + FechaIni + "),0,FMAR" + FechaIni + ")+IIF(IsNull(FABR" + FechaIni + "),0,FABR" + FechaIni + ")+IIF(IsNull(FMAY" + FechaIni + "),0,FMAY" + FechaIni + ")+IIF(IsNull(FJUN" + FechaIni + "),0,FJUN" + FechaIni + ")+IIF(IsNull(FJUL" + FechaIni + "),0,FJUL" + FechaIni + ")+IIF(IsNull(FAGO" + FechaIni + "),0,FAGO" + FechaIni + ")+IIF(IsNull(FSEP" + FechaIni + "),0,FSEP" + FechaIni + ")+IIF(IsNull(FOCT" + FechaIni + "),0,FOCT" + FechaIni + ")+IIF(IsNull(FNOV" + FechaIni + "),0,FNOV" + FechaIni + ")+IIF(IsNull(FDIC" + FechaIni + "),0,FDIC" + FechaIni + @"), 
                    TOTFe" + FechaFin+ " = IIF(IsNull(FENE" + FechaFin + "),0,FENE" + FechaFin + ")+IIF(IsNull(FFEB" + FechaFin + "),0,FFEB" + FechaFin + ")+IIF(IsNull(FMAR" + FechaFin + "),0,FMAR" + FechaFin + ")+IIF(IsNull(FABR" + FechaFin + "),0,FABR" + FechaFin + ")+IIF(IsNull(FMAY" + FechaFin + "),0,FMAY" + FechaFin + ")+IIF(IsNull(FJUN" + FechaFin + "),0,FJUN" + FechaFin + ")+IIF(IsNull(FJUL" + FechaFin + "),0,FJUL" + FechaFin + ")+IIF(IsNull(FAGO" + FechaFin + "),0,FAGO" + FechaFin + ")+IIF(IsNull(FSEP" + FechaFin + "),0,FSEP" + FechaFin + ")+IIF(IsNull(FOCT" + FechaFin + "),0,FOCT" + FechaFin + ")+IIF(IsNull(FNOV" + FechaFin + "),0,FNOV" + FechaFin + ")+IIF(IsNull(FDIC" + FechaFin + "),0,FDIC" + FechaFin + @"),
                    TOTFe = IIF(IsNull(FENE" + FechaIni + "),0,FENE" + FechaIni + ")+IIF(IsNull(FFEB" + FechaIni + "),0,FFEB" + FechaIni + ")+IIF(IsNull(FMAR" + FechaIni + "),0,FMAR" + FechaIni + ")+IIF(IsNull(FABR" + FechaIni + "),0,FABR" + FechaIni + ")+IIF(IsNull(FMAY" + FechaIni + "),0,FMAY" + FechaIni + ")+IIF(IsNull(FJUN" + FechaIni + "),0,FJUN" + FechaIni + ")+IIF(IsNull(FJUL" + FechaIni + "),0,FJUL" + FechaIni + ")+IIF(IsNull(FAGO" + FechaIni + "),0,FAGO" + FechaIni + ")+IIF(IsNull(FSEP" + FechaIni + "),0,FSEP" + FechaIni + ")+IIF(IsNull(FOCT" + FechaIni + "),0,FOCT" + FechaIni + ")+IIF(IsNull(FNOV" + FechaIni + "),0,FNOV" + FechaIni + ")+IIF(IsNull(FDIC" + FechaIni + "),0,FDIC" + FechaIni + ") + IIF(IsNull(FENE" + FechaFin + "),0,FENE" + FechaFin + ")+IIF(IsNull(FFEB" + FechaFin + "),0,FFEB" + FechaFin + ")+IIF(IsNull(FMAR" + FechaFin + "),0,FMAR" + FechaFin + ")+IIF(IsNull(FABR" + FechaFin + "),0,FABR" + FechaFin + ")+IIF(IsNull(FMAY" + FechaFin + "),0,FMAY" + FechaFin + ")+IIF(IsNull(FJUN" + FechaFin + "),0,FJUN" + FechaFin + ")+IIF(IsNull(FJUL" + FechaFin + "),0,FJUL" + FechaFin + ")+IIF(IsNull(FAGO" + FechaFin + "),0,FAGO" + FechaFin + ")+IIF(IsNull(FSEP" + FechaFin + "),0,FSEP" + FechaFin + ")+IIF(IsNull(FOCT" + FechaFin + "),0,FOCT" + FechaFin + ")+IIF(IsNull(FNOV" + FechaFin + "),0,FNOV" + FechaFin + ")+IIF(IsNull(FDIC" + FechaFin + "),0,FDIC" + FechaFin + ")  WHERE IdEjecutores = " + idEjecutor + " AND IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: SUMA TOTAL DEVENGADO EJECUTORES CALENDARIZADOS 18 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor19(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+ sNombre + " SET ActividadEspecifica = '*' WHERE ActividadEspecifica is null AND IdEjecutores = " + idEjecutor + " AND IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA ACTIVIDAD ESPECIFICA NULA EJECUTORES CALENDARIZADOS 19 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }



        public void CalendarizarEjecutor20(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombre)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+ sNombre + " SET NoEQUIPO = '*', NOMBREEQUIPO = '*', ESTRUC = '*', RENTA = '*', OBJETIVO = '*', INTERVALO = '*', PROF = '*', Notas = '*', Fec_Inic = IIF(IsNull(Fec_Inic),'01/01/2016',Fec_Inic), Fec_Term = IIF(IsNull(Fec_Term),'31/12/2016',Fec_Term), Días = IIF(IsNull([Días]),365,[Días]) WHERE pozo in ('Tareas de Apoyo','MO') OR NoEquipo is null AND IdEjecutores = " + idEjecutor + " AND IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA EJECUTORES CALENDARIZADOS 20 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor21(decimal idEjecutor, decimal idPOM,decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombrePOM+" SET DENE"+FechaIni+" = IIF(IsNull(DENE"+FechaIni+ "),0,DENE" + FechaIni + "), DFEB" + FechaIni+" = IIF(IsNull(DFEB"+FechaIni+ "),0,DFEB" + FechaIni + "), DMAR" + FechaIni+" = IIF(IsNull(DMAR"+FechaIni+ "),0,DMAR" + FechaIni + "), DABR" + FechaIni+" = IIF(IsNull(DABR"+FechaIni+ "),0,DABR" + FechaIni + "), DMAY" + FechaIni+" = IIF(IsNull(DMAY"+FechaIni+ "),0,DMAY" + FechaIni + "), DJUN" + FechaIni+" = IIF(IsNull(DJUN"+FechaIni+ "),0,DJUN" + FechaIni + "), DJUL" + FechaIni+" = IIF(IsNull(DJUL"+FechaIni+ "),0,DJUL" + FechaIni + "), DAGO" + FechaIni+" = IIF(IsNull(DAGO"+FechaIni+ "),0,DAGO" + FechaIni + "), DSEP" + FechaIni+" = IIF(IsNull(DSEP"+FechaIni+ "),0,DSEP" + FechaIni + "), DOCT" + FechaIni+" = IIF(IsNull(DOCT"+FechaIni+ "),0,DOCT" + FechaIni + "), DNOV" + FechaIni+" = IIF(IsNull(DNOV"+FechaIni+ "),0,DNOV" + FechaIni + "), DDIC" + FechaIni+" = IIF(IsNull(DDIC"+FechaIni+ "),0,DDIC" + FechaIni + "), FEne" + FechaIni+" = IIF(IsNull(FENE"+FechaIni+ "),0,FENE" + FechaIni + "), FFeb" + FechaIni+" = IIF(IsNull(FFEB"+FechaIni+ "),0,FFEB" + FechaIni + "), FMar" + FechaIni+" = IIF(IsNull(FMAR"+FechaIni+ "),0,FMAR" + FechaIni + "), FAbr" + FechaIni+" = IIF(IsNull(FABR"+FechaIni+ "),0,FABR" + FechaIni + "), FMay" + FechaIni+" = IIF(IsNull(FMAY"+FechaIni+ "),0,FMAY" + FechaIni + "), FJun" + FechaIni+" = IIF(IsNull(FJUN"+FechaIni+ "),0,FJUN" + FechaIni + "), FJul" + FechaIni+" = IIF(IsNull(FJUL"+FechaIni+ "),0,FJUL" + FechaIni + "), FAgo" + FechaIni+" = IIF(IsNull(FAGO"+FechaIni+ "),0,FAGO" + FechaIni + "), FSep" + FechaIni+" = IIF(IsNull(FSEP"+FechaIni+ "),0,FSEP" + FechaIni + "), FOct" + FechaIni+" = IIF(IsNull(FOCT"+FechaIni+ "),0,FOCT" + FechaIni + "), FNov" + FechaIni+" = IIF(IsNull(FNOV"+FechaIni+ "),0,FNOV" + FechaIni + "), FDic" + FechaIni+" = IIF(IsNull(FDIC"+FechaIni+ "),0,FDIC" + FechaIni + "), FENE" + FechaFin+" = IIF(IsNull(FENE"+FechaFin+ "),0,FENE" + FechaFin + "), FFEB" + FechaFin+" = IIF(IsNull(FFEB"+FechaFin+ "),0,FFEB" + FechaFin + "), FMAR" + FechaFin+" = IIF(IsNull(FMAR"+FechaFin+ "),0,FMAR" + FechaFin + "), FABR" + FechaFin+" = IIF(IsNull(FABR"+FechaFin+ "),0,FABR" + FechaFin + "), FMAY" + FechaFin+" = IIF(IsNull(FMAY"+FechaFin+ "),0,FMAY" + FechaFin + "), FJUN" + FechaFin+" = IIF(IsNull(FJUN"+FechaFin+ "),0,FJUN" + FechaFin + "), FJUL" + FechaFin+" = IIF(IsNull(FJUL"+FechaFin+ "),0,FJUL" + FechaFin + "), FAGO" + FechaFin+" = IIF(IsNull(FAGO"+FechaFin+ "),0,FAGO" + FechaFin + "), FSEP" + FechaFin+" = IIF(IsNull(FSEP"+FechaFin+ "),0,FSEP" + FechaFin + "), FOCT" + FechaFin+" = IIF(IsNull(FOCT"+FechaFin+ "),0,FOCT" + FechaFin + "), FNOV" + FechaFin+" = IIF(IsNull(FNOV"+FechaFin+ "),0,FNOV" + FechaFin + "), FDIC" + FechaFin+" = IIF(IsNull(FDIC"+FechaFin+ "),0,FDIC" + FechaFin + ") WHERE IdEjecutores = " + idEjecutor + " AND IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA EJECUTORES CALENDARIZADOS 21 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarEjecutor22(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombrePOM+" INNER JOIN CatSupervisores ON "+sNombrePOM+".FICHA = CatSupervisores.FICHA SET "+sNombrePOM+ ".USUARIO = [CatSupervisores].[USUARIO] WHERE "+sNombrePOM+".IdEjecutores = " + idEjecutor + " AND "+sNombrePOM+".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA EJECUTORES CALENDARIZADOS 22 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }

        public void CalendarizarEjecutor23(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE "+sNombrePOM+" SET año = mid(["+sNombrePOM+ "].PROPRE,12,5) WHERE statusoriginal='M.O.' AND " + sNombrePOM + ".IdEjecutores = " + idEjecutor + " AND " + sNombrePOM + ".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA EJECUTORES CALENDARIZADOS 23 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }



        public void CalendarizarEjecutor24(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE CatAdmon INNER JOIN "+sNombrePOM+" ON CatAdmon.Llave=["+sNombrePOM+"].Año SET "+sNombrePOM+".Cve_Campo = CatAdmon.[Cve_Campo], "+sNombrePOM+".Id_Asignacion = CatAdmon.[Id_Asignacion], "+sNombrePOM+".DescripcionAsignacion = CatAdmon.[Des_Asignacion], "+sNombrePOM+".Cve_Asignacion = CatAdmon.Cve_Asignacion WHERE ["+sNombrePOM+ "].statusoriginal='M.O.' AND "+sNombrePOM+".IdEjecutores = " + idEjecutor + " AND "+sNombrePOM+".IdPOM = " + idPOM;

                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA EJECUTORES CALENDARIZADOS 24 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }



        public void CalendarizarEjecutor25(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE CatAdmon INNER JOIN "+sNombrePOM+" ON CatAdmon.Llave="+sNombrePOM+".Año SET "+sNombrePOM+".Cve_Campo = CatAdmon.[Cve_Campo], "+sNombrePOM+".Id_Asignacion = CatAdmon.[Id_Asignacion], "+sNombrePOM+".DescripcionAsignacion = CatAdmon.[Des_Asignacion], "+sNombrePOM+".Cve_Asignacion = CatAdmon.Cve_Asignacion WHERE "+sNombrePOM+ ".[statusoriginal]='M.O.' AND " + sNombrePOM + ".IdEjecutores = " + idEjecutor + " AND " + sNombrePOM + ".IdPOM = " + idPOM;


                comando = new OleDbCommand(q, con);

                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA EJECUTORES CALENDARIZADOS 25 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }

        public void CalendarizarEjecutor26(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                //INSERTA DATOS TABLA EJECUTORES CALENDARIZADOS
                q = @"UPDATE CatIdPozo INNER JOIN "+sNombrePOM+" ON CatIdPozo.Pozo = "+sNombrePOM+".POZO SET "+sNombrePOM+".IIP = CatIdPozo.IIP";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();



                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA EJECUTORES CALENDARIZADOS 26 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }


        public void CalendarizarPolitica180(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                q = @"UPDATE "+ sNombrePOM + " SET fJul"+ FechaIni +" = dEne"+FechaIni+" WHERE politicaPago='180' and dEne" + FechaIni +" <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fAgo" + FechaIni + " = dFeb" + FechaIni + " WHERE politicaPago='180' and dFeb" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fSep" + FechaIni + " = dMar" + FechaIni + " WHERE politicaPago='180' and dMar" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fOct" + FechaIni + " = dAbr" + FechaIni + " WHERE politicaPago='180' and dAbr" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fNov" + FechaIni + " = dMay" + FechaIni + " WHERE politicaPago='180' and dMay" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fDic" + FechaIni + " = dJun" + FechaIni + " WHERE politicaPago='180' and dJun" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fEne" + FechaFin + " = dJul" + FechaIni + " WHERE politicaPago='180' and dJul" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fFeb" + FechaFin + " = dAgo" + FechaIni + " WHERE politicaPago='180' and dAgo" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fMar" + FechaFin + " = dSep" + FechaIni + " WHERE politicaPago='180' and dSep" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fAbr" + FechaFin + " = dOct" + FechaIni + " WHERE politicaPago='180' and dOct" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fMay" + FechaFin + " = dNov" + FechaIni + " WHERE politicaPago='180' and dNov" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJun" + FechaFin + " = dDic" + FechaIni + " WHERE politicaPago='180' and dDic" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA POLITICA DE PAGO 180 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }



        public void CalendarizarPolitica120(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                q = @"UPDATE " + sNombrePOM + " SET fMay" + FechaIni + " = dEne" + FechaIni + " WHERE politicaPago='120' and dEne" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fJun" + FechaIni + " = dFeb" + FechaIni + " WHERE politicaPago='120' and dFeb" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJul" + FechaIni + " = dMar" + FechaIni + " WHERE politicaPago='120' and dMar" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fAgo" + FechaIni + " = dAbr" + FechaIni + " WHERE politicaPago='120' and dAbr" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fSep" + FechaIni + " = dMay" + FechaIni + " WHERE politicaPago='120' and dMay" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fOct" + FechaIni + " = dJun" + FechaIni + " WHERE politicaPago='120' and dJun" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fNov" + FechaIni + " = dJul" + FechaIni + " WHERE politicaPago='120' and dJul" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fDic" + FechaIni + " = dAgo" + FechaIni + " WHERE politicaPago='120' and dAgo" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fEne" + FechaFin + " = dSep" + FechaIni + " WHERE politicaPago='120' and dSep" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fFeb" + FechaFin + " = dOct" + FechaIni + " WHERE politicaPago='120' and dOct" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fMar" + FechaFin + " = dNov" + FechaIni + " WHERE politicaPago='120' and dNov" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fAbr" + FechaFin + " = dDic" + FechaIni + " WHERE politicaPago='120' and dDic" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA POLITICA DE PAGO 120 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }

        }


        public void CalendarizarPolitica90(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                q = @"UPDATE " + sNombrePOM + " SET fAbr" + FechaIni + " = dEne" + FechaIni + " WHERE politicaPago='90' and dEne" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fMay" + FechaIni + " = dFeb" + FechaIni + " WHERE politicaPago='90' and dFeb" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJun" + FechaIni + " = dMar" + FechaIni + " WHERE politicaPago='90' and dMar" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJul" + FechaIni + " = dAbr" + FechaIni + " WHERE politicaPago='90' and dAbr" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fAgo" + FechaIni + " = dMay" + FechaIni + " WHERE politicaPago='90' and dMay" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fSep" + FechaIni + " = dJun" + FechaIni + " WHERE politicaPago='90' and dJun" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fOct" + FechaIni + " = dJul" + FechaIni + " WHERE politicaPago='90' and dJul" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fNov" + FechaIni + " = dAgo" + FechaIni + " WHERE politicaPago='90' and dAgo" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fDic" + FechaIni + " = dSep" + FechaIni + " WHERE politicaPago='90' and dSep" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fEne" + FechaFin + " = dOct" + FechaIni + " WHERE politicaPago='90' and dOct" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fFeb" + FechaFin + " = dNov" + FechaIni + " WHERE politicaPago='90' and dNov" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fMar" + FechaFin + " = dDic" + FechaIni + " WHERE politicaPago='90' and dDic" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA POLITICA DE PAGO 90 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }

        }


        public void CalendarizarPolitica60(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                q = @"UPDATE " + sNombrePOM + " SET fMar" + FechaIni + " = dEne" + FechaIni + " WHERE politicaPago='60' and dEne" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fAbr" + FechaIni + " = dFeb" + FechaIni + " WHERE politicaPago='60' and dFeb" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fMay" + FechaIni + " = dMar" + FechaIni + " WHERE politicaPago='60' and dMar" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJun" + FechaIni + " = dAbr" + FechaIni + " WHERE politicaPago='60' and dAbr" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJul" + FechaIni + " = dMay" + FechaIni + " WHERE politicaPago='60' and dMay" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fAgo" + FechaIni + " = dJun" + FechaIni + " WHERE politicaPago='60' and dJun" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fSep" + FechaIni + " = dJul" + FechaIni + " WHERE politicaPago='60' and dJul" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fOct" + FechaIni + " = dAgo" + FechaIni + " WHERE politicaPago='60' and dAgo" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fNov" + FechaIni + " = dSep" + FechaIni + " WHERE politicaPago='60' and dSep" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fDic" + FechaIni + " = dOct" + FechaIni + " WHERE politicaPago='60' and dOct" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fEne" + FechaFin + " = dNov" + FechaIni + " WHERE politicaPago='60' and dNov" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fFeb" + FechaFin + " = dDic" + FechaIni + " WHERE politicaPago='60' and dDic" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA POLITICA DE PAGO 60 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }

        }


        public void CalendarizarPolitica45(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                q = @"UPDATE " + sNombrePOM + " SET fMar" + FechaIni + " = dEne" + FechaIni + " WHERE politicaPago='45' and dEne" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fAbr" + FechaIni + " = dFeb" + FechaIni + " WHERE politicaPago='45' and dFeb" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fMay" + FechaIni + " = dMar" + FechaIni + " WHERE politicaPago='45' and dMar" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJun" + FechaIni + " = dAbr" + FechaIni + " WHERE politicaPago='45' and dAbr" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJul" + FechaIni + " = dMay" + FechaIni + " WHERE politicaPago='45' and dMay" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fAgo" + FechaIni + " = dJun" + FechaIni + " WHERE politicaPago='45' and dJun" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fSep" + FechaIni + " = dJul" + FechaIni + " WHERE politicaPago='45' and dJul" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fOct" + FechaIni + " = dAgo" + FechaIni + " WHERE politicaPago='45' and dAgo" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fNov" + FechaIni + " = dSep" + FechaIni + " WHERE politicaPago='45' and dSep" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fDic" + FechaIni + " = dOct" + FechaIni + " WHERE politicaPago='45' and dOct" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fEne" + FechaFin + " = dNov" + FechaIni + " WHERE politicaPago='45' and dNov" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fFeb" + FechaFin + " = dDic" + FechaIni + " WHERE politicaPago='45' and dDic" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA POLITICA DE PAGO 45 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }

        }



        public void CalendarizarPolitica2030(decimal idEjecutor, decimal idPOM, decimal FechaIni, decimal FechaFin, string sNombrePOM)
        {
            try
            {


                string q = "";
                OleDbConnection con = new OleDbConnection();
                OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();



                q = @"UPDATE " + sNombrePOM + " SET fFeb" + FechaIni + " = dEne" + FechaIni + " WHERE politicaPago in ('20', '30') and dEne" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fMar" + FechaIni + " = dFeb" + FechaIni + " WHERE politicaPago in ('20', '30') and dFeb" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fAbr" + FechaIni + " = dMar" + FechaIni + " WHERE politicaPago in ('20', '30') and dMar" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fMay" + FechaIni + " = dAbr" + FechaIni + " WHERE politicaPago in ('20', '30') and dAbr" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJun" + FechaIni + " = dMay" + FechaIni + " WHERE politicaPago in ('20', '30') and dMay" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fJul" + FechaIni + " = dJun" + FechaIni + " WHERE politicaPago in ('20', '30') and dJun" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();


                q = @"UPDATE " + sNombrePOM + " SET fAgo" + FechaIni + " = dJul" + FechaIni + " WHERE politicaPago in ('20', '30') and dJul" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fSep" + FechaIni + " = dAgo" + FechaIni + " WHERE politicaPago in ('20', '30') and dAgo" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fOct" + FechaIni + " = dSep" + FechaIni + " WHERE politicaPago in ('20', '30') and dSep" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fNov" + FechaIni + " = dOct" + FechaIni + " WHERE politicaPago in ('20', '30') and dOct" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fDic" + FechaIni + " = dNov" + FechaIni + " WHERE politicaPago in ('20', '30') and dNov" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                q = @"UPDATE " + sNombrePOM + " SET fEne" + FechaFin + " = dDic" + FechaIni + " WHERE politicaPago in ('20', '30') and dDic" + FechaIni + " <>0;";
                comando = new OleDbCommand(q, con);
                comando.ExecuteNonQuery();

                con.Close();



            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "ERROR: ACTUALIZA POLITICA DE PAGO 2030 " + ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }

        }

        #endregion

        #region BUSCARPOM

        public DataSet BuscarPOM(string sFechaIni, string sFechaFin)
        {
            try
            {
                string query = "SELECT Id,Nombre,Fecha FROM CatPOM where DateValue(Fecha) BETWEEN DateValue('" + sFechaIni + "') and DateValue('" + sFechaFin + "')";

                OleDbParameter[] opParametros = { };

                return fSQL(query, opParametros);

            }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

            }
            return null;
         
        }

        #endregion

        #region REQUERIDO
        public DataSet ObtenerRequerido(string sNombreTabla)
        {
            try
            {
                string query = "SELECT * FROM "+sNombreTabla+" ORDER BY idEjecutores DESC";

                OleDbParameter[] opParametros = { };

                return fSQL(query, opParametros);

            }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

            }
            return null;

        }
        #endregion

        #region UPDATE REQUERIDO
        public void UpdateRequerido(string sNombreTabla, string id, string column, string value)
        {
            try
            {
                string q = "";
                OleDbConnection con = new OleDbConnection();
                //OleDbCommand comando;
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();


                q = @"UPDATE " + sNombreTabla + @" set " +
                 column +   @" = '" + value + @"' " +
                 @"where id=" + id;
                //string query = "SELECT * FROM " + sNombreTabla + " ORDER BY idEjecutores DESC";

                OleDbCommand comando = new OleDbCommand(q, con);

                comando.Parameters.Clear();
               
                comando.ExecuteNonQuery();

                con.Close();

            }
            catch (Exception ex)
            {
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

            }
        }
        #endregion

        #region AJUSTE
        public DataSet CrearTablaAjusteporPozo(decimal idPOM, string nombreTabla,decimal FechaIni, decimal FechaFin)
        {
            string query = @"SELECT " + nombreTabla + @".POZO, Sum(" + nombreTabla + @".DENE2017+" + nombreTabla + @".DFEB2017+" + nombreTabla + @".DMAR2017+" + nombreTabla + @".DABR2017+" + nombreTabla + @".DMAY2017+" + nombreTabla + @".DJUN2017+" + nombreTabla + @".DJUL2017+" + nombreTabla + @".DAGO2017+" + nombreTabla + @".DSEP2017+" + nombreTabla + @".DOCT2017+" + nombreTabla + @".DNOV2017+" + nombreTabla + @".DDIC2017) AS ImpPozo, CostoPozo.Total, CostoPozo.Total/Sum(" + nombreTabla + @".DENE2017+" + nombreTabla + @".DFEB2017+" + nombreTabla + @".DMAR2017+" + nombreTabla + @".DABR2017+" + nombreTabla + @".DMAY2017+" + nombreTabla + @".DJUN2017+" + nombreTabla + @".DJUL2017+" + nombreTabla + @".DAGO2017+" + nombreTabla + @".DSEP2017+" + nombreTabla + @".DOCT2017+" + nombreTabla + @".DNOV2017+" + nombreTabla + @".DDIC2017) AS Division
                FROM " + nombreTabla + @" INNER JOIN CostoPozo ON " + nombreTabla + @".POZO = CostoPozo.Pozo
                GROUP BY " + nombreTabla + @".POZO, CostoPozo.Total";

            OleDbParameter[] opParametros = { };

            return fSQL(query, opParametros);


            
        }


        public DataSet CrearTablaAjustePozoSuma(decimal idPOM, string nombreTabla, decimal FechaIni, decimal FechaFin)
        {
            string query = @"SELECT *
FROM (select Id, count(*) as Registros," + nombreTabla + @".pozo  from " + nombreTabla + @" INNER JOIN PorcentajeAjustePozo ON " + nombreTabla + @".POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DEne2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DFeb2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DMar2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DAbr2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DMay2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DJun2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DJul2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DAgo2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DSep2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DOct2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DNov2017 > 0
group by Id, EjecutoresCalendarizados.pozo
Union
select Id, count(*),EjecutoresCalendarizados.pozo  from EjecutoresCalendarizados INNER JOIN PorcentajeAjustePozo ON EjecutoresCalendarizados.POZO = PorcentajeAjustePozo.POZO
where PorcentajeAjustePozo.Division > 1
and DDic2017 > 0
group by Id, EjecutoresCalendarizados.pozo )  AS [%$##@_Alias];
";

            OleDbParameter[] opParametros = { };

            return fSQL(query, opParametros);



        }





        #endregion

        #region Insert New Row
        public void GuardarCatNewRow(DataRowView row, string catalogo)
        {
            try
            {

                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\BaseDeDatos.accdb";

                con.Open();

                if (catalogo == "Actividades")
                {
                    string q = "INSERT INTO CatActividades (PRO_PRE,POZO,ACTIVIDAD,ACTIVIDAD_ESPECIFICA,No_EQUIPO,NOMBRE_DE_EQUIPO,ESTRUC,RENTA,OBJETIVO,INTERVALO,PROF,Fec_Inic,Fec_Term,Días,NOTAS,Campo16,LLAVE,LLAVE1,LLAVE2,LLAVE3,Campo21,ELEM_PEP,LLAVE_DE_CRUCE) VALUES (@pre,@pozo,@ACTIVIDAD,@ACTIVIDAD_ESPECIFICA,@No_EQUIPO,@NOMBRE_DE_EQUIPO,@ESTRUC,@RENTA,@OBJETIVO,@INTERVALO,@PROF,@Fec_Inic,@Fec_Term,@Días,@NOTAS,@Campo16,@LLAVE,@LLAVE1,@LLAVE2,@LLAVE3,@Campo21,@ELEM_PEP,@LLAVE_DE_CRUCE)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    DateTime fecha_i = Convert.ToDateTime(row[11].ToString());
                    DateTime fecha_f = Convert.ToDateTime(row[12].ToString());

                    comando.Parameters.AddWithValue("@pre", row[0].ToString());
                    comando.Parameters.AddWithValue("@pozo", row[1].ToString());
                    comando.Parameters.AddWithValue("@ACTIVIDAD", row[2].ToString());
                    comando.Parameters.AddWithValue("@ACTIVIDAD_ESPECIFICA", row[3].ToString());
                    comando.Parameters.AddWithValue("@No_EQUIPO", row[4].ToString());
                    comando.Parameters.AddWithValue("@NOMBRE_DE_EQUIPO", row[5].ToString());
                    comando.Parameters.AddWithValue("@ESTRUC", row[6].ToString());
                    comando.Parameters.AddWithValue("@RENTA", row[7].ToString());
                    comando.Parameters.AddWithValue("@OBJETIVO", row[8].ToString());
                    comando.Parameters.AddWithValue("@INTERVALO", row[9].ToString());
                    comando.Parameters.AddWithValue("@PROF", row[10].ToString());
                    comando.Parameters.AddWithValue("@Fec_Inic", fecha_i);
                    comando.Parameters.AddWithValue("@Fec_Term", fecha_f);
                    comando.Parameters.AddWithValue("@Días", row[13].ToString());
                    comando.Parameters.AddWithValue("@Campo16", row[14].ToString());
                    comando.Parameters.AddWithValue("@NOTAS", row[15].ToString());
                    comando.Parameters.AddWithValue("@LLAVE", row[16].ToString());
                    comando.Parameters.AddWithValue("@LLAVE1", row[17].ToString());
                    comando.Parameters.AddWithValue("@LLAVE2", row[18].ToString());
                    comando.Parameters.AddWithValue("@LLAVE3", row[19].ToString());
                    comando.Parameters.AddWithValue("@Campo21", row[20].ToString());
                    comando.Parameters.AddWithValue("@ELEM_PEP", row[21].ToString());
                    comando.Parameters.AddWithValue("@LLAVE_DE_CRUCE", row[22].ToString());

                    comando.ExecuteNonQuery();



                }
                else if (catalogo == "Admon")
                {
                    string q = "INSERT INTO CatAdmon (Llave,Cve_Campo,ID_Asignacion,Des_Asignacion,Cve_Asignacion,Des_Campo,Cve_Proyecto,Des_Proyecto,Campo_Cab_Asig,Activo,Proy_Prog) VALUES (@Llave,@Cve_Campo,@ID_Asignacion,@Des_Asignacion,@Cve_Asignacion,@Des_Campo,@Cve_Proyecto,@Des_Proyecto,@Campo_Cab_Asig,@Activo,@Proy_Prog)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    comando.Parameters.AddWithValue("@Llave", row[0].ToString());
                    comando.Parameters.AddWithValue("@Cve_Campo", row[1].ToString());
                    comando.Parameters.AddWithValue("@ID_Asignacion", row[2].ToString());
                    comando.Parameters.AddWithValue("@Des_Asignacion", row[3].ToString());
                    comando.Parameters.AddWithValue("@Cve_Asignacion", row[4].ToString());
                    comando.Parameters.AddWithValue("@Des_Campo", row[5].ToString());
                    comando.Parameters.AddWithValue("@Cve_Proyecto", row[6].ToString());
                    comando.Parameters.AddWithValue("@Des_Proyecto", row[7].ToString());
                    comando.Parameters.AddWithValue("@Campo_Cab_Asig", row[8].ToString());
                    comando.Parameters.AddWithValue("@Activo", row[9].ToString());
                    comando.Parameters.AddWithValue("@Proy_Prog", row[10].ToString());

                    comando.ExecuteNonQuery();

                }
                else if (catalogo == "AsigCampoPYIN")
                {
                    string q = "INSERT INTO CatAsigCampoPYIN (Campo_Cab_Asig,Subdirección,ID_Asignacion,Cve_Asignacion,Des_Asignacion,Cve_Campo,Des_Campo,Cve_Proyecto,Des_Proyecto,Campo10,Activo,Estatus,F_inicio,F_fin) VALUES (@Campo_Cab_Asig,@Subdirección,@ID_Asignacion,@Cve_Asignacion,@Des_Asignacion,@Cve_Campo,@Des_Campo,@Cve_Proyecto,@Des_Proyecto,@Campo10,@Activo,@Estatus,@F_inicio,@F_fin)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    DateTime fecha_i = Convert.ToDateTime(row[12].ToString());
                    DateTime fecha_f = Convert.ToDateTime(row[13].ToString());

                    comando.Parameters.AddWithValue("@Campo_Cab_Asig", row[0].ToString());
                    comando.Parameters.AddWithValue("@Subdirección", row[1].ToString());
                    comando.Parameters.AddWithValue("@ID_Asignacion", row[2].ToString());
                    comando.Parameters.AddWithValue("@Cve_Asignacion", row[3].ToString());
                    comando.Parameters.AddWithValue("@Des_Asignacion", row[4].ToString());

                    comando.Parameters.AddWithValue("@Cve_Campo", row[5].ToString());
                    comando.Parameters.AddWithValue("@Des_Campo", row[6].ToString());
                    comando.Parameters.AddWithValue("@Cve_Proyecto", row[7].ToString());
                    comando.Parameters.AddWithValue("@Des_Proyecto", row[8].ToString());
                    comando.Parameters.AddWithValue("@Campo10", row[9].ToString());
                    comando.Parameters.AddWithValue("@Activo", row[10].ToString());
                    comando.Parameters.AddWithValue("@Estatus", row[11].ToString());
                    comando.Parameters.AddWithValue("@F_inicio", fecha_i);
                    comando.Parameters.AddWithValue("@F_fin", fecha_f);

                    comando.ExecuteNonQuery();

                }
                else if (catalogo == "Contratos")
                {
                    string q = "INSERT INTO CatContratos (CONTRATO,RESERVA,Descripción_Contrato,FecIncContrato,FecVencContrato,POLITICAPAGO,COMPAÑÍA,PROVEEDOR,CONCATENAR) VALUES (@CONTRATO,@RESERVA,@Descripción_Contrato,@FecIncCont,@FecVencContrato,@POLITICAPAGO,@COMPAÑÍA,@PROVEEDOR,@CONCATENAR)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    //Agregado porque causa conflicto si el campo esta vacion porque es numerico en la base de datos
                    var pol = row[5].ToString() == "" ? "0" : row[5].ToString();
                    var pro = row[7].ToString() == "" ? "0" : row[7].ToString();

                    comando.Parameters.AddWithValue("@CONTRATO", row[0].ToString());
                    comando.Parameters.AddWithValue("@RESERVA", row[1].ToString());
                    comando.Parameters.AddWithValue("@Descripción_Contrato", row[2].ToString());
                    comando.Parameters.AddWithValue("@FecIncCont", row[3].ToString());
                    comando.Parameters.AddWithValue("@FecVencContrato", row[4].ToString());
                    comando.Parameters.AddWithValue("@POLITICAPAGO", pol);
                    comando.Parameters.AddWithValue("@COMPAÑÍA", row[6].ToString());
                    comando.Parameters.AddWithValue("@PROVEEDOR", pro);
                    comando.Parameters.AddWithValue("@CONCATENAR", row[8].ToString());

                    comando.ExecuteNonQuery();

                }
                else if (catalogo == "IdPozo")
                {
                    string q = "INSERT INTO CatIdPozo (Id,Pozo,IIP) VALUES (@Id,@Pozo,@IIP)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    Decimal id = row[0].ToString() == "" ? 0 : Convert.ToDecimal(row[0].ToString());

                    comando.Parameters.AddWithValue("@Id", id);
                    comando.Parameters.AddWithValue("@Pozo", row[1].ToString());
                    comando.Parameters.AddWithValue("@IIP", row[2].ToString());

                    comando.ExecuteNonQuery();

                }
                else if (catalogo == "LlaveControl")
                {
                    string q = "INSERT INTO CatLlaveControl (Id1,prott,PP,id_presentacion,ID,ID2) VALUES (@Id1,@prott,@PP,@id_presentacion,@ID,@ID2)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    Decimal id1 = row[0].ToString() == "" ? 0 : Convert.ToDecimal(row[0].ToString());

                    comando.Parameters.AddWithValue("@Id1", id1);
                    comando.Parameters.AddWithValue("@prott", row[1].ToString());
                    comando.Parameters.AddWithValue("@PP", row[2].ToString());
                    comando.Parameters.AddWithValue("@id_presentacion", row[3].ToString());
                    comando.Parameters.AddWithValue("@ID", row[4].ToString());
                    comando.Parameters.AddWithValue("@ID2", row[5].ToString());

                    comando.ExecuteNonQuery();

                }
                else if (catalogo == "PPEquivalentes")
                {
                    string q = "INSERT INTO CatPPEquivalentes (Programa_anterior,Elemento_PEP_Anterior,Descripción_PP_Anterior,Programa_Vigente,Elemento_PEP,Descripción_PP) VALUES (@Programa_anterior,@Elemento_PEP_Anterior,@Descripción_PP_Anterior,@Programa_Vigente,@Elemento_PEP,@Descripción_PP)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    comando.Parameters.AddWithValue("@Programa_anterior", row[0].ToString());
                    comando.Parameters.AddWithValue("@Elemento_PEP_Anterior", row[1].ToString());
                    comando.Parameters.AddWithValue("@Descripción_PP_Anterior", row[2].ToString());
                    comando.Parameters.AddWithValue("@Programa_Vigente", row[3].ToString());
                    comando.Parameters.AddWithValue("@Elemento_PEP", row[4].ToString());
                    comando.Parameters.AddWithValue("@Descripción_PP", row[5].ToString());
                    comando.ExecuteNonQuery();
                }
                else if (catalogo == "Proyectos")
                {
                    string q = "INSERT INTO CatProyectos (PROYECTO,PYIN) VALUES (@PROYECTO,@PYIN)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    comando.Parameters.AddWithValue("@PROYECTO", row[0].ToString());
                    comando.Parameters.AddWithValue("@PYIN", row[1].ToString());
                    comando.ExecuteNonQuery();

                }
                else if (catalogo == "POZOADMON")
                {
                    string q = " INSERT INTO CatPOZOADMON (PROGRAMA, ELEMENTO_PEP, PRO_POZO, CONCATENAR, POZO, Soc_CO, Entidad_CP, Proy_Aval, PROY, EPEP_Descripcion, EPEP_Ind, EPEP_Cent, Prog_Inic, Prog_Term, Fecha, Prog, Tox, Campo,PTO) VALUES(@PROGRAMA, @ELEMENTO_PEP, @PRO_POZO, @CONCATENAR, @POZO, @Soc_CO, @Entidad_CP, @Proy_Aval, @PROY, @EPEP_Descripcion, @EPEP_Ind, @EPEP_Cent, @Prog_Inic, @Prog_Term, @Fecha, @Prog, @Tox, @Campo, @PTO)";

                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    Decimal cent = row[11].ToString() == "" ? 0 : Convert.ToDecimal(row[11].ToString());
                    DateTime fi = Convert.ToDateTime(row[12].ToString()); // Convert.ToDateTime(row[11].ToString());
                    DateTime ff = Convert.ToDateTime(row[13].ToString()); // Convert.ToDateTime(row[12].ToString());

                    comando.Parameters.AddWithValue("@PROGRAMA", row[0].ToString());
                    comando.Parameters.AddWithValue("@ELEMENTO_PEP", row[1].ToString());
                    comando.Parameters.AddWithValue("@PRO_POZO", row[2].ToString());
                    comando.Parameters.AddWithValue("@CONCATENAR", row[3].ToString());
                    comando.Parameters.AddWithValue("@POZO", row[4].ToString());
                    comando.Parameters.AddWithValue("@Soc_CO", row[5].ToString());
                    comando.Parameters.AddWithValue("@Entidad_CP", row[6].ToString());
                    comando.Parameters.AddWithValue("@Proy_Aval", row[7].ToString());
                    comando.Parameters.AddWithValue("@PROY", row[8].ToString());
                    comando.Parameters.AddWithValue("@EPEP_Descripcion", row[9].ToString());
                    comando.Parameters.AddWithValue("@EPEP_Ind", row[10].ToString());
                    comando.Parameters.AddWithValue("@EPEP_Cent", cent);
                    comando.Parameters.AddWithValue("@Prog_Inic", fi);
                    comando.Parameters.AddWithValue("@Prog_Term", ff);
                    comando.Parameters.AddWithValue("@Fecha", row[14].ToString());
                    comando.Parameters.AddWithValue("@Prog", row[15].ToString());
                    comando.Parameters.AddWithValue("@Tox", row[16].ToString());
                    comando.Parameters.AddWithValue("@Campo", row[17].ToString());
                    comando.Parameters.AddWithValue("@PTO", row[18].ToString());

                    comando.ExecuteNonQuery();

                }
                else if (catalogo == "Supervisores")
                {
                    string q = "INSERT INTO CatSupervisores (USUARIO,FICHA) VALUES (@USUARIO,@FICHA)";
                    OleDbCommand comando = new OleDbCommand(q, con);

                    comando.Parameters.Clear();

                    comando.Parameters.AddWithValue("@USUARIO", row[0].ToString());
                    comando.Parameters.AddWithValue("@FICHA", row[1].ToString());
                    comando.ExecuteNonQuery();

                }

                con.Close();
                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = "El Cat" + catalogo + " fué actualizado con exito.";

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);

            }
            catch (Exception ex)
            {

                var alert = new RadDesktopAlert();
                alert.Header = "NOTIFICACIÓN";
                alert.Content = ex;
                alert.CanAutoClose = false;

                RadDesktopAlertManager manager = new RadDesktopAlertManager();
                manager.ShowAlert(alert);
            }



        }
        #endregion


    }

}





