    using System;
using System.Data;
using System.Web;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using System.IO;
using JobSiteStarterKit.BOL;
using JobSiteStarterKit.DAL;
using Flanders.Library.WebControls.Resources;
using System.Web.Script.Serialization;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using System.Linq;
using Org.BouncyCastle.Utilities;
using System.Globalization;
using DevExpress.Web.ASPxEditors;
using System.Text.RegularExpressions;
using DevExpress.Web.ASPxGridView;
using ExcelDataReader;
using DevExpress.Web.ASPxUploadControl;
using OfficeOpenXml;
using System.Configuration;

public partial class CargaMasivaFacturas_aspx : System.Web.UI.Page
{


    private static string[] arrayArchivos;
    private static DataTable DaTa = new DataTable();
    private static DataRow row;

    private static string path;
    private static string errorExcel;
    private static int numErrorExcel;
    private static int numerrorfactura;
    private static int numokfactura;

    byte[] fichero_contenido;

    public class Payments_Contratos
    {
        public int id { get; set; }
        public string deal_code { get; set; }
        public string name { get; set; }
        public double amount_usd { get; set; }
        public double amount_eur { get; set; }
    }
    public class Token_Inversiones
    {
        public string x_api_key { get; set; }
    }

    List<Payments_Contratos> contracts;


    protected void Page_Load(object sender, EventArgs e)
    {

        try
        {

            ASPxButtonProcesar.Visible = false;

            if (!IsPostBack)
            {
               
              
            }
        }
        catch (Exception ex)
        {
            JobSiteStarterKit.DAL.traza.TrazaPBASE.WriteError(ex.Message, ex.Source);
        }

    }

  
    protected void ASPxButtonDescargaPlantilla_Click(object sender, EventArgs e)
    {

        Response.Redirect("./CargaMasivafacturas.aspx");
    }

    public void CrearDataTable()
    {
        // Variables para las columnas y las filas
        DataColumn column;

        if(DaTa.Columns.Count ==0)
        {
           
            // Se tiene que crear primero la columna asignandole Nombre y Tipo de datos    
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "tipoproveedor";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "empresaid";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "empresa";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "providerid";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "providername";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "CIF";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "CCC";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "SWIFT";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "IBAN";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "ABA";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AccountNumber";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "BankName";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "projectcode";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "numfactura";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "fechafactura";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "concepto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "moneda";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "formapago";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "cuentaorigen";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "importebruto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "tipoiva";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "tasa";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "IRPF";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "fechaorevistapago";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "idtipogasto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "tipogasto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "iddirectoindirecto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "directoindirecto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "idgrupo";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "grupo";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "clavegrupo";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "idsubtipogasto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "subtipogasto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "clavesubtipogasto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "cccta";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "iddepto";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "departamento";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "estado";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "nombredocumento";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "entidadfinanciadora";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "codigocontrato";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "bankaccountid";
            DaTa.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "noSujeto";
            DaTa.Columns.Add(column);
        }

        DaTa.Clear();
    }

    //Rellena una columna con un string
    public ExcelWorksheet FillCol(ExcelWorksheet ws, int filaInicial, char letraColumna, string[] data)
    {

        for (int fila = filaInicial; fila < data.Length + 2; fila++)
        {
            ws.Cells[letraColumna + fila.ToString()].Value = data[fila - 2];
        }


        return ws;
    }

    private ExcelWorksheet GenerateMatrix(string[] ListaTipos, string[] ColumnaListaTipos, string[] ColumnaDatos, ExcelWorksheet ws, string tablaPrincipal)
    {
        int numTablas = 0;
        string[] arrayTablas = new string[ListaTipos.Length];

        OfficeOpenXml.Table.ExcelTableCollection tblcollection = ws.Tables;

        for (int i = 0; i < ListaTipos.Length; i++)
        {
            int longitud = 0;

            for (int j = 0; j < ColumnaDatos.Length; j++)
            {
                if (ListaTipos[i] == ColumnaListaTipos[j])
                {
                    ws.Cells[longitud + 2, i + 3, longitud + 2, i + 3].Value = ColumnaDatos[j];
                    longitud++;
                }
            }

            if (longitud != 0)
            {
                arrayTablas[i] = tablaPrincipal + "_" + (numTablas).ToString();

                using (ExcelRange Rng = ws.Cells[1, i + 3, longitud + 1, i + 3])
                {
                    OfficeOpenXml.Table.ExcelTable table = tblcollection.Add(Rng, tablaPrincipal + "_" + (numTablas).ToString());
                }
                numTablas++;
            }

        }

        for (int i = 0; i < arrayTablas.Length; i++)
        {
            ws.Cells[i + 2, 1, i + 2, 1].Value = ListaTipos[i];
            ws.Cells[i + 2, 2, i + 2, 2].Value = arrayTablas[i];
        }

        //ws.Cells["C31"].Value = "ARRIBA RANGO LOCAL";

        //var val = ws.DataValidations.AddListValidation("C30");
        //val.Formula.ExcelFormula = "=$A$2:$A$" + (numTablas + 2);

        //var val2 = ws.DataValidations.AddListValidation("D30");
        //val2.Formula.ExcelFormula = "=INDIRECT(VLOOKUP(C30,A2:B" + (arrayTablas.Length + 1).ToString() + ",2,FALSE))";

        return ws;
    }

    public void GeneraExcel(object sender, EventArgs e)
    {
        // DEFINIMOS UN LÍMITE RAZONABLE PARA EL USUARIO
        // 5000 filas es más que suficiente para una carga manual y evita que Excel se congele al abrir.
        const int MAX_ROWS = 100;

        byte[] content;
        using (ExcelPackage package = new ExcelPackage(new FileInfo(HttpContext.Current.Server.MapPath("~/plantillas/") + "MassiveUploadFacturasTemplates.xlsx")))
        {
            content = package.GetAsByteArray();
        }
        MemoryStream mem = new MemoryStream(content);
        OfficeOpenXml.ExcelPackage pck = new OfficeOpenXml.ExcelPackage(mem);

        #region Strings de datos

        string[] strIdEmpresa, strNameEmpresa;
        cargarEmpresa(out strIdEmpresa, out strNameEmpresa);

        string[] strTipoProveedor = cargarTipoProveedor();
        string[] strTipoProveedorProveedor, strProveedor1;
        DataTable dtProveedor = Providers.GetProvidersActiveAll();
        cargarTipoProveedorProveedor(out strTipoProveedorProveedor, out strProveedor1, dtProveedor);

        string[] strProveedor = cargarProveedor(dtProveedor);
        string[] strProveedorFormaPago, strFormaPago;
        System.Data.DataTable dtProveedorActivo = Providers.GetProvidersActiveAll();
        cargarProveedorFormaPago(out strProveedorFormaPago, out strFormaPago, dtProveedorActivo);

        string[] strProveedorTasa, strTasa;
        cargarProveedorTasa(out strProveedorTasa, out strTasa, dtProveedorActivo);

        string[] strProyecto = cargarProyecto();
        string[] strMoneda = cargaMoneda();

        string[] strCompany = cargarEmpresa1();
        string[] strCompanyId, strBankAccount;
        cargarCuentaOrigen(out strCompanyId, out strBankAccount);

        string[] strTipoGasto = cargarTipoGasto();
        string[] strTipoGastoDirectoIndirecto, strDirectoIndirecto;
        cargarDirectoIndirecto(out strTipoGastoDirectoIndirecto, out strDirectoIndirecto);

        string[] strGrupo = cargarDirectoIndirecto1();
        string[] strGrupoGasto, strClave;
        cargarGruposGastosPagosManuales(out strGrupoGasto, out strClave);

        string[] strGrupo1 = cargarGrupo1();
        string[] strGrupo2, strClave2;
        cargarSubGrupo(out strGrupo2, out strClave2);

        string[] strEmpresa = cargarEmpresa1();
        string[] strEmpresaDepartamento, strDepartamento;
        cargarEmpresaDepartamento(out strEmpresaDepartamento, out strDepartamento);

        DataTable dtContrato = LoadContracts();
        string[] strContratos = cargarContratos(dtContrato);

        //Cuentas Bancarias
        DataTable dtCuentasBancarias = Providers.GetCuentasBancariasProveedoresAll();
        string[] strCuentaProveedor, strCuenta;
        cargarCuantasBancarias(out strCuentaProveedor, out strCuenta, dtCuentasBancarias);

        // GASTOS POR PROVEEDOR/EMPRESA
        DataTable dtProviderGasto = new DataTable();
        dtProviderGasto.Columns.Add("ID");
        dtProviderGasto.Columns.Add("CompanyID");
        dtProviderGasto.Columns.Add("ProviderID");
        dtProviderGasto.Columns.Add("ProjectID");
        dtProviderGasto.Columns.Add("NumFactura");
        dtProviderGasto.Columns.Add("Importe");
        dtProviderGasto.Columns.Add("Tipo_Gasto");
        dtProviderGasto.Columns.Add("TipoDireccion");
        dtProviderGasto.Columns.Add("Grupo");
        dtProviderGasto.Columns.Add("Subtipo");
        dtProviderGasto.Columns.Add("Departamento");
        dtProviderGasto.Columns.Add("Observaciones");

        try
        {
            string providerIdFilter = string.Empty;
            string companyIdFilter = string.Empty;

            if (Request != null)
            {
                providerIdFilter = (Request["ProviderID"] ?? Request["providerid"] ?? string.Empty).Trim();
                companyIdFilter = (Request["CompanyID"] ?? Request["companyid"] ?? Request["Company"] ?? Request["company"] ?? string.Empty).Trim();
            }

            string connectionString = string.Empty;
            string[] connectionCandidates = { "PBASE", "PBase", "ConnectionString", "DefaultConnection", "JobSiteStarterKitConnectionString" };

            foreach (string candidateName in connectionCandidates)
            {
                ConnectionStringSettings candidate = ConfigurationManager.ConnectionStrings[candidateName];
                if (candidate != null && !string.IsNullOrWhiteSpace(candidate.ConnectionString))
                {
                    connectionString = candidate.ConnectionString;
                    break;
                }
            }

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                foreach (ConnectionStringSettings connection in ConfigurationManager.ConnectionStrings)
                {
                    if (connection == null || string.IsNullOrWhiteSpace(connection.ConnectionString))
                    {
                        continue;
                    }

                    if (connection.Name == "LocalSqlServer")
                    {
                        continue;
                    }

                    connectionString = connection.ConnectionString;
                    break;
                }
            }

            if (!string.IsNullOrWhiteSpace(connectionString))
            {
                StringBuilder query = new StringBuilder();
                query.Append("SELECT ");
                query.Append("pg.[ID], pg.[CompanyID], pg.[ProviderID], pg.[ProjectID], pg.[NumFactura], ");
                query.Append("pg.[Importe], pg.[Tipo_Gasto], pg.[TipoDireccion], pg.[Grupo], pg.[Subtipo], ");
                query.Append("pg.[Departamento], pg.[Observaciones] ");
                query.Append("FROM [dbo].[tbl_Provider_Gasto] AS pg ");
                query.Append("WHERE pg.[Active] = 1 ");

                if (!string.IsNullOrWhiteSpace(providerIdFilter))
                {
                    query.Append("AND pg.[ProviderID] = @ProviderID ");
                }

                if (!string.IsNullOrWhiteSpace(companyIdFilter))
                {
                    query.Append("AND (pg.[CompanyID] = @CompanyID OR pg.[CompanyID] IS NULL) ");
                }

                using (System.Data.SqlClient.SqlConnection sqlConnection = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    using (System.Data.SqlClient.SqlCommand sqlCommand = new System.Data.SqlClient.SqlCommand(query.ToString(), sqlConnection))
                    {
                        if (!string.IsNullOrWhiteSpace(providerIdFilter))
                        {
                            sqlCommand.Parameters.AddWithValue("@ProviderID", providerIdFilter);
                        }

                        if (!string.IsNullOrWhiteSpace(companyIdFilter))
                        {
                            sqlCommand.Parameters.AddWithValue("@CompanyID", companyIdFilter);
                        }

                        using (System.Data.SqlClient.SqlDataAdapter sqlAdapter = new System.Data.SqlClient.SqlDataAdapter(sqlCommand))
                        {
                            dtProviderGasto.Rows.Clear();
                            sqlAdapter.Fill(dtProviderGasto);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            JobSiteStarterKit.DAL.traza.TrazaPBASE.WriteError(ex.Message, ex.Source);
        }

        #endregion

        ExcelWorksheet wsEmpresa = pck.Workbook.Worksheets["Empresa"];
        wsEmpresa = FillCol(wsEmpresa, 2, 'B', strIdEmpresa);
        wsEmpresa = FillCol(wsEmpresa, 2, 'C', strNameEmpresa);
        wsEmpresa = FillCol(wsEmpresa, 2, 'D', strIdEmpresa);

        ExcelWorksheet wsProveedor = pck.Workbook.Worksheets["Proveedor"];
        wsProveedor = FillCol(wsProveedor, 2, 'B', strProveedor);
        wsProveedor = FillCol(wsProveedor, 2, 'C', strFormaPago);

        ExcelWorksheet wsProyecto = pck.Workbook.Worksheets["Proyecto"];
        wsProyecto = FillCol(wsProyecto, 2, 'B', strProyecto);

        ExcelWorksheet wsMoneda = pck.Workbook.Worksheets["Moneda"];
        wsMoneda = FillCol(wsMoneda, 2, 'B', strMoneda);

        ExcelWorksheet wsTipoProveedorProveedor = pck.Workbook.Worksheets["TipoProveedor_and_Proveedor"];
        wsTipoProveedorProveedor = GenerateMatrix(strTipoProveedor, strTipoProveedorProveedor, strProveedor1, wsTipoProveedorProveedor, "TablaTipoProveedorProveedor");

        ExcelWorksheet wsProveedorFormaPago = pck.Workbook.Worksheets["Proveedor_and_FormaPago"];
        wsProveedorFormaPago = GenerateMatrix(strProveedor, strProveedorFormaPago, strFormaPago, wsProveedorFormaPago, "TablaProveedorFormaPago");

        ExcelWorksheet wsProveedorTasa = pck.Workbook.Worksheets["Proveedor_and_Tasa"];
        wsProveedorTasa = GenerateMatrix(strProveedor, strProveedorTasa, strTasa, wsProveedorTasa, "TablaProveedorTasa");

        ExcelWorksheet wsEmpresaCuentaOrigen = pck.Workbook.Worksheets["Empresa_and_CuentaOrigen"];
        wsEmpresaCuentaOrigen = GenerateMatrix(strCompany, strCompanyId, strBankAccount, wsEmpresaCuentaOrigen, "TablaEmpresaCuentaOrigen");

        ExcelWorksheet wsTipoGastoDirectoIndirecto = pck.Workbook.Worksheets["TipoGasto_and_DirectoIndirecto"];
        wsTipoGastoDirectoIndirecto = GenerateMatrix(strTipoGasto, strTipoGastoDirectoIndirecto, strDirectoIndirecto, wsTipoGastoDirectoIndirecto, "TablaTipoGastoDirectoIndirecto");

        ExcelWorksheet wsGrupo = pck.Workbook.Worksheets["DirectoIndirecto_and_Grupo"];
        wsGrupo = GenerateMatrix(strGrupo, strGrupoGasto, strClave, wsGrupo, "TablaDirectoIndirectoGrupo");

        ExcelWorksheet wsSubGrupo = pck.Workbook.Worksheets["Grupo_and_SubGrupo"];
        wsSubGrupo = GenerateMatrix(strGrupo1, strGrupo2, strClave2, wsSubGrupo, "TablaGrupoSubGrupo");

        ExcelWorksheet wsEmpresaDepartamento = pck.Workbook.Worksheets["Empresa_and_Departamento"];
        wsEmpresaDepartamento = GenerateMatrix(strEmpresa, strEmpresaDepartamento, strDepartamento, wsEmpresaDepartamento, "TablaEmpresaDepartamento");

        ExcelWorksheet wsContrato = pck.Workbook.Worksheets["Contrato"];
        wsEmpresa = FillCol(wsContrato, 2, 'B', strContratos);

        ExcelWorksheet wsEmpresaEmpresa = pck.Workbook.Worksheets["Empresa_and_Empresa"];
        wsEmpresaEmpresa = GenerateMatrix(strNameEmpresa, strNameEmpresa, strIdEmpresa, wsEmpresaEmpresa, "TablaEmpresaEmpresa");

        ExcelWorksheet wsIban = pck.Workbook.Worksheets["Iban"];
        wsIban = GenerateMatrix(strProveedor, strCuentaProveedor, strCuenta, wsIban, "TablaIban");

        ExcelWorksheet wsProveedorGasto = pck.Workbook.Worksheets["Proveedor_Gasto"];
        if (wsProveedorGasto == null)
        {
            wsProveedorGasto = pck.Workbook.Worksheets.Add("Proveedor_Gasto");
        }
        else
        {
            wsProveedorGasto.Cells.Clear();
        }
        wsProveedorGasto.Cells[1, 1].LoadFromDataTable(dtProviderGasto, true);

        #region FILL MAIN WS

        ExcelWorksheet mainWs = pck.Workbook.Worksheets["Facturas"];

        //TipoProveedor
        mainWs.DataValidations.Clear();
        var range = ExcelRange.GetAddress(2, 1, MAX_ROWS, 1);
        var val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=TipoProveedor!$B$2:$B$3";

        //Empresa
        mainWs.DataValidations.Clear();
        range = ExcelRange.GetAddress(2, 2, MAX_ROWS, 2);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=Empresa!$C$2:$C$" + (strIdEmpresa.Length + 1).ToString();

        //Proveedor
        range = ExcelRange.GetAddress(2, 3, MAX_ROWS, 3);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($A2,TipoProveedor_and_Proveedor!$A$2:$B$" + (strProveedor.Length + 1).ToString() + ",2,FALSE))";

        //Proyecto
        mainWs.DataValidations.Clear();
        range = ExcelRange.GetAddress(2, 4, MAX_ROWS, 4);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=Proyecto!$B$2:$B$" + (strProyecto.Length + 1).ToString();

        //Moneda
        mainWs.DataValidations.Clear();
        range = ExcelRange.GetAddress(2, 8, MAX_ROWS, 8);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=Moneda!$B$2:$B$" + (strMoneda.Length + 1).ToString();

        //FormaPago
        range = ExcelRange.GetAddress(2, 9, MAX_ROWS, 9);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($C2,Proveedor_and_FormaPago!$A$2:$B$" + (strProveedor.Length + 1).ToString() + ",2,FALSE))";

        //CuentaOrigen
        range = ExcelRange.GetAddress(2, 10, MAX_ROWS, 10);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($Y2,Empresa_and_CuentaOrigen!$A$2:$B$" + (strCompany.Length + 1).ToString() + ",2,FALSE))";

        //TipoIVA
        mainWs.DataValidations.Clear();
        range = ExcelRange.GetAddress(2, 12, MAX_ROWS, 12);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=TipoIVA!$B$2:$B$3";

        //Tasa
        range = ExcelRange.GetAddress(2, 13, MAX_ROWS, 13);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($C2,Proveedor_and_Tasa!$A$2:$B$" + (strProveedor.Length + 1).ToString() + ",2,FALSE))";

        //IRPF
        range = ExcelRange.GetAddress(2, 14, MAX_ROWS, 14);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=IRPF!$B$2:$B$11";

        //TIPO GASTO
        range = ExcelRange.GetAddress(2, 16, MAX_ROWS, 16);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=TipoGasto!$B$2:$B$7";

        //DirectoIndirecto
        range = ExcelRange.GetAddress(2, 17, MAX_ROWS, 17);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($P2,TipoGasto_and_DirectoIndirecto!$A$2:$B$" + (strCompany.Length + 1).ToString() + ",2,FALSE))";

        //Grupo
        range = ExcelRange.GetAddress(2, 18, MAX_ROWS, 18);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($Q2,DirectoIndirecto_and_Grupo!$A$2:$B$" + (strCompany.Length + 1).ToString() + ",2,FALSE))";

        //Subtipo
        range = ExcelRange.GetAddress(2, 19, MAX_ROWS, 19);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($R2,Grupo_and_SubGrupo!$A$2:$B$" + (strGrupo1.Length + 1).ToString() + ",2,FALSE))";

        //Departamento
        range = ExcelRange.GetAddress(2, 20, MAX_ROWS, 20);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($Y2,Empresa_and_Departamento!$A$2:$B$" + (strGrupo1.Length + 1).ToString() + ",2,FALSE))";

        //Estado
        range = ExcelRange.GetAddress(2, 21, MAX_ROWS, 21);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=Estado!$B$2:$B$4";

        //Entidad Financiadora
        range = ExcelRange.GetAddress(2, 23, MAX_ROWS, 23);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=Entidad!$B$2:$B$3";

        //Contratos
        if (strContratos.Length != 0)
        {
            range = ExcelRange.GetAddress(2, 24, MAX_ROWS, 24);
            val = mainWs.DataValidations.AddListValidation(range);
            val.Formula.ExcelFormula = "=Contrato!$B$2:$B$" + (strContratos.Length + 1).ToString();
        }
        
        if (strIdEmpresa.Length > 0)
        {
            string rangoBusqueda = "Empresa!$C$2:$D$" + (strIdEmpresa.Length + 1).ToString();
          
            for (int i = 2; i <= MAX_ROWS; i++)
            {                
                mainWs.Cells[i, 25].Formula = "IF(ISBLANK(B" + i + "), \"\", VLOOKUP(B" + i + "," + rangoBusqueda + ",2,FALSE))";
            }
        }

        mainWs.Column(25).Width = 0;

        //IBAN
        range = ExcelRange.GetAddress(2, 27, MAX_ROWS, 27);
        val = mainWs.DataValidations.AddListValidation(range);
        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP($C2,Iban!$A$2:$B$" + (strProveedor.Length + 1).ToString() + ",2,FALSE))";

        #endregion

        wsEmpresa.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsProveedor.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsProyecto.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsMoneda.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsEmpresaCuentaOrigen.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsTipoGastoDirectoIndirecto.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsGrupo.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsSubGrupo.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsEmpresaDepartamento.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsProveedorFormaPago.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsProveedorTasa.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsEmpresaEmpresa.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsTipoProveedorProveedor.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsIban.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        wsProveedorGasto.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;

        ExcelWorksheet wsEstado = pck.Workbook.Worksheets["Estado"];
        wsEstado.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;

        ExcelWorksheet wsEntidad = pck.Workbook.Worksheets["Entidad"];
        wsEntidad.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;

        ExcelWorksheet wsTipoProveedor = pck.Workbook.Worksheets["TipoProveedor"];
        wsTipoProveedor.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;

        ExcelWorksheet wsTipoIVA = pck.Workbook.Worksheets["TipoIVA"];
        wsTipoIVA.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;

        ExcelWorksheet wsIRPF = pck.Workbook.Worksheets["IRPF"];
        wsIRPF.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;

        ExcelWorksheet wsTipoGasto = pck.Workbook.Worksheets["TipoGasto"];
        wsTipoGasto.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;

        ExcelWorksheet wsCuentaOrigen = pck.Workbook.Worksheets["CuentaOrigen"];
        wsCuentaOrigen.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;

        pck.Workbook.CalcMode = ExcelCalcMode.Manual;

        MemoryStream save = new MemoryStream();
        pck.SaveAs(save);
        byte[] bn = save.ToArray();

        Response.Clear();

        HttpCookie cookie = new HttpCookie("downloadStatus", "completed");
        cookie.HttpOnly = false;
        cookie.Path = "/";
        Response.Cookies.Add(cookie);

        Response.Clear();
        Response.AddHeader("Cache-Control", "no-cache, must-revalidate, post-check=0, pre-check=0");
        Response.AddHeader("Pragma", "no-cache");
        Response.AddHeader("Content-Description", "File Download");
        Response.AddHeader("Content-Type", "application/force-download");
        Response.AddHeader("Content-Transfer-Encoding", "binary\n");
        Response.AddHeader("content-disposition", "attachment;filename=PLANTILLA_CARGA_MASIVA.xlsx");
        Response.BinaryWrite(bn);
        Response.End();
    }


    public void RellenaFormulaAux(int col, int fil, int longArray, string s1, string s2, string s3, ExcelWorksheet wsAUX)
    {


        for (int i = fil; i < 1000; i++)
        {
            wsAUX.Cells[i, col, i, col].Formula = s1 + i.ToString() + s2 + longArray.ToString() + s3;
        }
    }

    protected void ASPxButtonValidar_Click(object sender, EventArgs e)
    {




        HttpFileCollection uploadedFiles = Request.Files;
        if (FileUpload.FileName != "" && myfile.PostedFile.ContentLength != 0)
        {
            string selectedFileName = Path.GetFileName(FileUpload.FileName);
            string fileExtension = System.IO.Path.GetExtension(selectedFileName).ToLower();
            if (fileExtension == ".xls" || fileExtension == ".xlsx")
            {
                if (FileUpload.FileName != "PLANTILLA_CARGA_MASIVA" && FileUpload.FileName != "PLANTILLA_CARGA_MASIVA.xls" && FileUpload.FileName != "PLANTILLA_CARGA_MASIVA.xlsx")
                {
                    ASPxLabelMsg.Text = "Por favor, el nombre de la plantilla subida debe ser el mismo que el de la descargada";
                    ASPxLabelMsg.Visible = true;
                    return;
                }

                //path = SOTEC.Utilities.TempUpload.SaveFileToTempDir(selectedFileName, FileUpload.FileBytes);
                AlmacenarTemp();
                string aux = HttpContext.Current.Server.MapPath("~/App_Data/tmp/subidamasiva/") + selectedFileName;
                File.Delete(aux);
                SOTEC.Utilities.TempUpload.DeleteFileFromTempDir(selectedFileName);
            }
            else
            {
                ASPxLabelMsg.Text = "Por favor, selecciona un fichero Excel.";
                ASPxLabelMsg.Visible = true;
                return;
            }


        }
        else
        {
            ASPxLabelMsg.Text = "Ambos archivos (plantilla y ficheros) se deben seleccionar.";
            ASPxLabelMsg.Visible = true;
            return;
        }


    }

    protected void AlmacenarTemp()
    {
        HttpFileCollection uploadedFiles = Request.Files;
        arrayArchivos = new string[uploadedFiles.Count];

        string selectedFileName = Path.GetFileName(FileUpload.FileName);
        path = SOTEC.Utilities.TempUpload.SaveFileToTempDir(selectedFileName, FileUpload.FileBytes);


        var realPath = HttpContext.Current.Server.MapPath("~/App_Data/tmp/subidamasiva/");
        if (!Directory.Exists(realPath))
        {
            Directory.CreateDirectory(realPath);
        }






        for (int i = 0; i < uploadedFiles.Count; i++)
        {
            HttpPostedFile userPostedFile = uploadedFiles[i];

            if (userPostedFile.ContentLength > 0)
            {
                string aux = realPath + "\\" + userPostedFile.FileName;
                userPostedFile.SaveAs(aux);
                arrayArchivos[i] = userPostedFile.FileName;
            }

        }

        loadExcel(path, arrayArchivos);


        SOTEC.Utilities.TempUpload.DeleteFileFromTempDir(selectedFileName);





    }

    #region CARGAS LISTAS
    private void cargarCuantasBancarias( out string[] str1, out string[] str2, DataTable dt)
    {
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[0].ToString().Trim().ToUpper();
            str2[i] = row[6].ToString().Trim().ToUpper();
            i++;
        }

    }
    private DataTable LoadContracts()
    {

        // Creamos la tabla
        DataTable dtContracts = new DataTable();

        try
        {
            string token = getUser();
            contracts = getContracts(token);

            

            System.Data.DataColumn c;

            c = new System.Data.DataColumn();
            c.ColumnName = "id";
            c.DataType = Type.GetType("System.Int32");
            dtContracts.Columns.Add(c);

            c = new System.Data.DataColumn();
            c.ColumnName = "deal_code";
            c.DataType = Type.GetType("System.String");
            dtContracts.Columns.Add(c);

            c = new System.Data.DataColumn();
            c.ColumnName = "name";
            c.DataType = Type.GetType("System.String");
            dtContracts.Columns.Add(c);

            c = new System.Data.DataColumn();
            c.ColumnName = "amount_usd";
            c.DataType = Type.GetType("System.Double");
            dtContracts.Columns.Add(c);

            c = new System.Data.DataColumn();
            c.ColumnName = "amount_eur";
            c.DataType = Type.GetType("System.Double");
            dtContracts.Columns.Add(c);

            DataTable dtPaymentsDesglose = Payments.getAllDesgloseInversion();

            bool encontrado = false;
            foreach (Payments_Contratos r in contracts)
            {
                if (!string.IsNullOrEmpty(r.deal_code))
                {
                    foreach (DataRow data in dtPaymentsDesglose.Rows)
                    {
                        if (data["ContractCode"].ToString() == r.deal_code)
                        {
                            encontrado = true;
                            break;
                        }
                    }

                    if (!encontrado)
                    {
                        System.Data.DataRow row = dtContracts.NewRow();
                        row["id"] = r.id;
                        row["deal_code"] = r.deal_code;
                        row["name"] = "[" + r.deal_code + "]: " + r.name;
                        row["amount_usd"] = r.amount_usd;
                        row["amount_eur"] = r.amount_eur;

                        dtContracts.Rows.Add(row);

                    }
                    encontrado = false;
                }

            }

           

        }
        catch { }

        return dtContracts;
    }

    private string[] cargaMoneda()
    {
        System.Data.DataTable dt = Legal.getAllCurrenciesAbreviado();

        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[1].ToString();
            i++;
        }

        return stringArr;

    }

    protected void cargarEmpresa(out string[] str1, out string[] str2)
    {

        System.Data.DataTable dt = Employee.GetCompanies();
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[0].ToString().Trim().ToUpper();
            str2[i] = row[1].ToString().Trim().ToUpper();
            i++;
        }
    }

    private string[] cargarContratos(DataTable dt)
    {
        
        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[2].ToString().Trim().ToUpper();
            i++;
        }

        return stringArr;

    }

	private string[] cargarTipoProveedor()
    {
        var stringArr = new string[2];

        stringArr[0] = "PROVEEDOR";
        stringArr[1] = "EMPLEADO";


        return stringArr;

    }

    private string[] cargarTipoGasto()
    {
        System.Data.DataTable dt = Payments.GetTiposGasto();

        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[1].ToString().Trim().ToUpper();
            i++;
        }

        return stringArr;

    }


    private string[] cargarEmpresa1()
    {
        System.Data.DataTable dt = Employee.GetCompanies();

        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[0].ToString().Trim().ToUpper();
            i++;
        }

        return stringArr;

    }

    private string[] cargarProveedor(DataTable dt)
    {
   

        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[6].ToString().Trim().ToUpper();
            i++;
        }

        return stringArr;

    }

    protected void cargarProveedorFormaPago(out string[] str1, out string[] str2,DataTable dt)
    {

     
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[6].ToString().Trim().ToUpper();
            str2[i] = row[12].ToString().Trim().ToUpper();
            i++;
        }
    }

    protected void cargarTipoProveedorProveedor(out string[] str1, out string[] str2,DataTable dt)
    {

  

        DataTable dtAuxProveedores = new DataTable();

        DataColumn c;

        c = new System.Data.DataColumn();
        c.ColumnName = "Type";
        c.DataType = Type.GetType("System.String");
        dtAuxProveedores.Columns.Add(c);

        c = new System.Data.DataColumn();
        c.ColumnName = "Name";
        c.DataType = Type.GetType("System.String");
        dtAuxProveedores.Columns.Add(c);

     

        HashSet<string> seenRows = new HashSet<string>();

        foreach (DataRow row in dt.Rows)
        {
            string name = row["NombreCombo"].ToString();
            string type = row["TipoProveedorCompleto"].ToString();
            string combined = name + "|" + type;

            if (!seenRows.Contains(combined))
            {
                DataRow newRow = dtAuxProveedores.NewRow();
                newRow["Name"] = name;
                newRow["Type"] = type;
                dtAuxProveedores.Rows.Add(newRow);
                seenRows.Add(combined);
            }
        }


        str1 = new string[dtAuxProveedores.Rows.Count];
        str2 = new string[dtAuxProveedores.Rows.Count];
        int i = 0;
        foreach (DataRow row in dtAuxProveedores.Rows)
        {
            str1[i] = row["Type"].ToString().Trim().ToUpper();
            str2[i] = row["Name"].ToString().Trim().ToUpper();
            i++;
        }
    }

    protected void cargarProveedorTasa(out string[] str1, out string[] str2,DataTable dt)
    {
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[6].ToString().Trim().ToUpper();
            str2[i] = "0.21";
            i++;
        }
    }


    protected void cargarCuentaOrigen(out string[] str1, out string[] str2)
    {

        System.Data.DataTable dt = Payments.GetCuentaOrigen("-1");
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[1].ToString().Trim().ToUpper();
            str2[i] = row[2].ToString().Trim().ToUpper();
            i++;
        }
    }

    protected void cargarGruposGastosPagosManuales(out string[] str1, out string[] str2)
    {

        System.Data.DataTable dt = Payments.GetGruposGastosManuales();
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[4].ToString().Trim().ToUpper();
            str2[i] = row[1].ToString().Trim().ToUpper();
            i++;
        }
    }

    protected void cargarDirectoIndirecto(out string[] str1, out string[] str2)
    {

        System.Data.DataTable dt = Payments.GetGastoDirectoIndirectoAll();
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[5].ToString().Trim().ToUpper();
            str2[i] = row[1].ToString().Trim().ToUpper();
            i++;
        }
    }

   
    private string[] cargarDirectoIndirecto1()
    {
        System.Data.DataTable dt = Payments.GetGastoDirectoIndirectoAll();

        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[1].ToString().Trim().ToUpper();
            i++;
        }

        return stringArr;

    }

    private string[] cargarDepartamento(string Empresa)
    {
        System.Data.DataTable dt = Departamentos.GetDepartamentosActive(Empresa);

        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[2].ToString().Trim().ToUpper();
            i++;
        }

        return stringArr;

    }

    private string[] cargarGrupo1()
    {
        System.Data.DataTable dt = Payments.GetGrupoSubgrupo();

        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[1].ToString().Trim().ToUpper();
            i++;
        }

        return stringArr;

    }

    protected void cargarSubGrupo(out string[] str1, out string[] str2)
    {

        System.Data.DataTable dt = Payments.GetGrupoSubgrupo();
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[1].ToString().Trim().ToUpper();
            str2[i] = row[4].ToString().Trim().ToUpper();
            i++;
        }
    }

    protected void cargarEmpresaDepartamento(out string[] str1, out string[] str2)
    {

        System.Data.DataTable dt = Departamentos.GetAllDepartamentosPorEmpresa();
        str1 = new string[dt.Rows.Count];
        str2 = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            str1[i] = row[2].ToString().Trim().ToUpper();
            str2[i] = row[3].ToString().Trim().ToUpper();
            i++;
        }
    }

    private string[] cargarProyecto()
    {
        System.Data.DataTable dt = Projects.GetProjects();

        var stringArr = new string[dt.Rows.Count];
        int i = 0;
        foreach (DataRow row in dt.Rows)
        {
            stringArr[i] = row[1].ToString();
            i++;
        }

        return stringArr;

    }

    #endregion


    protected void loadExcel(string selectedFileName, string[] arrayarchivos)
    {

        DataTable dtFacturas;
        CrearDataTable();


        var file = new FileInfo(selectedFileName);
        using (
            var stream = File.Open(selectedFileName, FileMode.Open, FileAccess.Read))
        {
            IExcelDataReader reader;

            if (file.Extension.Equals(".xls"))
                reader = ExcelDataReader.ExcelReaderFactory.CreateBinaryReader(stream, null);
            else if (file.Extension.Equals(".xlsx"))
                reader = ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(stream, null);
            else
                throw new Exception("Invalid FileName");

            //// reader.IsFirstRowAsColumnNames
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };

            var dataSet = reader.AsDataSet(conf);
            dtFacturas = dataSet.Tables[1];
        }


        //MARCAPÁGINAS


        #region lectura de cabeceras
        if (dtFacturas.Columns[0].ToString() != "Tipo Proveedor*" &&
        dtFacturas.Columns[1].ToString() != "Empresa*" &&
        dtFacturas.Columns[2].ToString() != "Proveedor*" &&
        dtFacturas.Columns[3].ToString() != "Proyecto*" &&
        dtFacturas.Columns[4].ToString() != "Nº Factura" &&
        dtFacturas.Columns[5].ToString() != "Fecha factura*" &&
        dtFacturas.Columns[6].ToString() != "Concepto" &&
        dtFacturas.Columns[7].ToString() != "Moneda*" &&
        dtFacturas.Columns[8].ToString() != "Forma Pago*" &&
        dtFacturas.Columns[9].ToString() != "Cuenta origen*" &&
        dtFacturas.Columns[10].ToString() != "Importe Bruto*" &&
        dtFacturas.Columns[11].ToString() != "Tipo IVA" &&
        dtFacturas.Columns[12].ToString() != "Tasa" &&
        dtFacturas.Columns[13].ToString() != "IRPF" &&
        dtFacturas.Columns[14].ToString() != "Fecha Prevista Pago*" &&
        dtFacturas.Columns[15].ToString() != "Tipo Gasto*" &&
        dtFacturas.Columns[16].ToString() != "Directo/Indirecto*" &&
        dtFacturas.Columns[17].ToString() != "Grupo*" &&
        dtFacturas.Columns[18].ToString() != "Subtipo gasto*" &&
        dtFacturas.Columns[19].ToString() != "Depto*" &&
        dtFacturas.Columns[23].ToString() != "Código de contrato" &&
        dtFacturas.Columns[24].ToString() != "No Sujeto" &&
        dtFacturas.Columns[26].ToString() != "Iban*"
        )

        {
            ASPxLabelMsg.Text = "La versión de la plantilla empleada no es la actual. Por favor, descargue y rellene la plantilla con los campos actuales";
            ASPxLabelMsg.Visible = true;
            return;
        }

        #endregion


        if (dtFacturas.Rows.Count == 0)
        {
            ASPxLabelMsg.Text = "No hay registros para tratar.";
            ASPxLabelMsg.Visible = true;
            return;
        }

        string errores = "";

        DataTable dtEmpresa = Employee.GetCompanies();
        DataTable dtProveedor = Providers.GetProvidersActive();
        DataTable dtEmpleado = Providers.GetProvidersActiveAll();
        DataTable dtProyectos = Projects.GetProjects();
        DataTable dtMonedas = Legal.getAllCurrenciesAbreviado();
        DataTable dtTipoGasto = Payments.GetTiposGasto();
        DataTable dtCuentaOrigen;
        DataTable dtDirectoIndirecto;
        DataTable dtGrupo;
        DataTable dtSubtipo;
        DataTable dtDepto;
        DataTable dtCCCta;
        DataTable dtDepartamento;
        DataTable dtEntidadFinanciadora = Payments.GetActiveEntidadFinanciadora();
        DataTable dtCodigoContrato = Payments.getAllDesgloseInversion();


        ASPxListBoxWarnings.Items.Clear();
        ASPxListBoxWarnings.Items.Add("Validacion de carga masiva de facturas");
        ASPxListBoxWarnings.Items.Add("------------------------------------------------");


        for (int r = 0; r < dtFacturas.Rows.Count; r++)
        {

            string aux = "";
            string auxprovidertype = "";
            string auxempresa = "";
            string auxproveedor = "";
            string auxproyecto = "";
            string auxmoneda = "";
            string auxformapago = "";
            string auxcuentaorigen = "";
            string auxtipogasto = "";
            string auxdirectoindirecto = "";
            string auxgrupo = "";
            string auxsubtipogasto = "";
            string auxdepto = "";
            string auxestado = "";
            string auxnombredocumento = "";
            string auxentidadfinanciada = "";
            string auxcodigocontrato = "";
            string auxnumfactura = "";
            bool compraapps = false;
            bool documentovalidado = true;

            numerrorfactura = 0;
            numokfactura = 0;


            //datos id
            string tipoproveedor = "";
            string empresaid = "";
            string empresa = "";


            string providerid = "";
            string providername = "";
            string formapago = "";
            string CIF = "";
            string CCC = "";
            string SWIFT = "";
            string IBAN = "";
            string ABA = "";
            string AccountNumber = "";
            string BankName = "";
            
            string projectcode = "";
            string numfactura = "";
            string fechafactura = "";
            string concepto = "";
            string moneda = "";
            string cuentaorigen = "";
            string importebruto = "";
            string tipoiva = "";
            string tasa = "";
            string IRPF = "";
            string fechaorevistapago = "";
            string idtipogasto = "";
            string tipogasto = "";
            string iddirectoindirecto = "";
            string directoindirecto = "";

            string idgrupo = "";
            string grupo = "";
            string clavegrupo = "";

            string idsubtipogasto = "";
            string subtipogasto = "";
            string clavesubtipogasto = "";

            string cccta = "";

            string iddepto = "";
            string departamento = "";
            string estado = "";
            string nombredocumento = "";
            string entidadfinanciadora = "";
            string codigocontrato = "";
            string bankaccountid = "";


            #region Faltan datos oligatorios






            if (dtFacturas.Rows[r][0].ToString() == "")
            {
                aux += "Tipo Proveedor, ";
            }
            if (dtFacturas.Rows[r][1].ToString() == "")
            {
                aux += "Empresa, ";
            }
            if (dtFacturas.Rows[r][2].ToString() == "")
            {
                aux += "Proveedor, ";
            }
            if (dtFacturas.Rows[r][3].ToString() == "")
            {
                aux += "Proyecto, ";
            }
            if (dtFacturas.Rows[r][4].ToString() == "")
            {
                aux += "N Factura, ";
            }
            if (dtFacturas.Rows[r][5].ToString() == "")
            {
                aux += "Fecha factura, ";
            }
            if (dtFacturas.Rows[r][7].ToString() == "")
            {
                aux += "Moneda, ";
            }
            if (dtFacturas.Rows[r][0].ToString() == "EMPLEADO")
            {

            }
            else {
                if (dtFacturas.Rows[r][8].ToString() == "")
                {
                    aux += "Forma Pago, ";
                } }
            if (dtFacturas.Rows[r][9].ToString() == "")
            {
                aux += "Cuenta origen, ";
            }
            if (dtFacturas.Rows[r][10].ToString() == "")
            {
                aux += "Importe bruto, ";
            }
            if (dtFacturas.Rows[r][14].ToString() == "")
            {
                aux += "Fecha Prevista Pago, ";
            }
            if (dtFacturas.Rows[r][15].ToString() == "")
            {
                aux += "Tipo Gasto, ";
            }
            if (dtFacturas.Rows[r][16].ToString() == "")
            {
                aux += "Directo/Indirecto, ";
            }
            if (dtFacturas.Rows[r][17].ToString() == "")
            {
                aux += "Grupo, ";
            }
            if (dtFacturas.Rows[r][18].ToString() == "")
            {
                aux += "Subtipo gasto, ";
            }
            if (dtFacturas.Rows[r][19].ToString() == "")
            {
                aux += "Depto, ";
            }
            if (dtFacturas.Rows[r][20].ToString() == "")
            {
                aux += "Estado, ";
            }
            if (dtFacturas.Rows[r][21].ToString() == "")
            {
                aux += "Nombre documento, ";
            }
            if (dtFacturas.Rows[r][26].ToString() == "")
            {
                aux += "Iban, ";
            }

            if (aux != "")
            {
                aux = aux.TrimEnd(' ');
                aux = aux.TrimEnd(',');
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: Faltan datos obligatorios en la fila: " + aux + ". ||";
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: Faltan datos obligatorios en la fila: " + aux + ".");
                numErrorExcel++;
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            string noSujetoData1 = dtFacturas.Rows[r][25].ToString();
            #region nombre deocumento con documentos cargados
            nombredocumento = dtFacturas.Rows[r][21].ToString();
            for (int i = 0; i < arrayArchivos.Length; i++)
            {
                if (arrayArchivos[i].ToString().ToUpper().Trim() == nombredocumento.ToUpper().Trim())
                {
                    documentovalidado = true;
                }
            }

            if (!documentovalidado)
            {
                auxnombredocumento = auxnombredocumento.TrimEnd(' ');
                auxnombredocumento = auxnombredocumento.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El documento " + dtFacturas.Rows[r][21].ToString().ToUpper().Trim() + " no fue cargado.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El documento " + dtFacturas.Rows[r][21].ToString().ToUpper().Trim() + " no fue cargado. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar Proveedor

            if (dtFacturas.Rows[r][0].ToString().ToUpper().Trim() != "PROVEEDOR" && dtFacturas.Rows[r][0].ToString().ToUpper().Trim() != "EMPLEADO")
            {
                auxprovidertype += dtFacturas.Rows[r][0].ToString().ToUpper().Trim();
            }
            else
            {
                if(dtFacturas.Rows[r][0].ToString().ToUpper().Trim()=="PROVEEDOR")
                {
                    tipoproveedor = "P";
                }
                if (dtFacturas.Rows[r][0].ToString().ToUpper().Trim() == "EMPLEADO")
                {
                    tipoproveedor = "E";
                }

            }

            if (auxprovidertype != "")
            {
                auxprovidertype = auxprovidertype.TrimEnd(' ');
                auxprovidertype = auxprovidertype.TrimEnd(',');
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El tipo de proveedor no tiene el formato correcto (PROVEEDOR/EMPLEADO): " + auxprovidertype + ". ||";
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El tipo de proveedor no tiene el formato correcto (PROVEEDOR/EMPLEADO): " + auxprovidertype + ".");
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar Empresa

            string filtro = "CompanyId = '" + dtFacturas.Rows[r][24].ToString().ToUpper() + "'";
            DataRow[] resultadoEmpresa = dtEmpresa.Select(filtro);
            
            if (resultadoEmpresa.Length == 0)
            {
                auxempresa += dtFacturas.Rows[r][1].ToString().ToUpper();
            }
            else
            {
                empresa = dtFacturas.Rows[r][1].ToString().ToUpper().Trim();
                empresaid = resultadoEmpresa[0]["CompanyId"].ToString().Trim();

                dtCuentaOrigen = Payments.GetCuentaOrigen(empresaid);
                bankaccountid= dtCuentaOrigen.Rows[0]["BankAccountId"].ToString().Trim(); 


            }

            if (auxempresa != "")
            {
                auxempresa = auxempresa.TrimEnd(' ');
                auxempresa = auxempresa.TrimEnd(',');
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: La empresa " + dtFacturas.Rows[r][1].ToString().ToUpper().Trim() + " no se encuentra dentro de las empresas. ||";
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La empresa " + dtFacturas.Rows[r][1].ToString().ToUpper().Trim() + " no se encuentra dentro de las empresas.");
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar Proveedor

            filtro = "NombreCombo = '" + dtFacturas.Rows[r][2].ToString().ToUpper().Trim() + "'";

            DataRow[] resultadoProveedor = dtProveedor.Select(filtro);
            DataRow[] resultadoEmpleado = dtEmpleado.Select(filtro);
            if (resultadoEmpleado.Length != 0)
            {
                if (resultadoEmpleado[0]["TipoProveedor"].ToString().Trim() == "E")
                {
                    providerid = resultadoEmpleado[0]["ProviderId"].ToString().Trim();
                    providername = resultadoEmpleado[0]["Nombre"].ToString().Trim();
                    CIF = resultadoEmpleado[0]["CIF"].ToString().Trim();
                    CCC = resultadoEmpleado[0]["CCC"].ToString().Trim();
                    formapago = resultadoEmpleado[0]["FormaPago"].ToString().Trim();
                    IBAN = resultadoEmpleado[0]["IBAN"].ToString().Trim();
                    CIF = resultadoEmpleado[0]["CIF"].ToString().Trim();
                    SWIFT = resultadoEmpleado[0]["SWIFT"].ToString().Trim();
                }
                else
                {
                    if (resultadoEmpleado[0]["TipoProveedor"].ToString().Trim() == "P") {
                        providerid = resultadoProveedor[0]["ProviderId"].ToString().Trim();
                        providername = resultadoProveedor[0]["Nombre"].ToString().Trim();
                        formapago = resultadoProveedor[0]["FormaPago"].ToString().Trim();
                        CIF = resultadoProveedor[0]["CIF"].ToString().Trim();


                        DataTable dtCuentasBancarias = Providers.GetCuentasBancariasProveedoresAll();
                       string filtroC = "IBAN = '" + dtFacturas.Rows[r][26].ToString().ToUpper().Trim() + "' AND ProviderID = '" + providerid + "'";

                        DataRow[] resultadoCuentas= dtCuentasBancarias.Select(filtroC);

                 
                        if (resultadoCuentas.Length!=0)
                        {
                            CCC = resultadoCuentas[0]["CCC"].ToString().Trim();
                            SWIFT = resultadoCuentas[0]["SWIFT"].ToString().Trim();
                            IBAN = resultadoCuentas[0]["IBAN"].ToString().Trim();
                            ABA = resultadoCuentas[0]["ABA"].ToString().Trim();
                            AccountNumber = resultadoCuentas[0]["AccountNumber"].ToString().Trim();
                            BankName = resultadoCuentas[0]["BankName"].ToString().Trim();
                        }
                        else
                        {

                            CCC = resultadoProveedor[0]["CCC"].ToString().Trim();
                            SWIFT = resultadoProveedor[0]["SWIFT"].ToString().Trim();
                            IBAN = resultadoProveedor[0]["IBAN"].ToString().Trim();
                            ABA = resultadoProveedor[0]["ABA"].ToString().Trim();
                            AccountNumber = resultadoProveedor[0]["AccountNumber"].ToString().Trim();
                            BankName = resultadoProveedor[0]["BankName"].ToString().Trim();

                        }
      
                    }
                    else {
                        
                    } 
                }
            }         
                
                else
                {
                    auxproveedor += dtFacturas.Rows[r][2].ToString().ToUpper();

            }
            
            

            if (auxproveedor != "")
            {
                auxproveedor = auxproveedor.TrimEnd(' ');
                auxproveedor = auxproveedor.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El proveedor/empleado " + dtFacturas.Rows[r][2].ToString().ToUpper().Trim() + " no se encuentra dentro de los proveedores/empleados activos.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El proveedor/empleado " + dtFacturas.Rows[r][2].ToString().ToUpper().Trim() + " no se encuentra dentro de los proveedores/empleados activos. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar Proyecto

            filtro = "ProjectTitle = '" + dtFacturas.Rows[r][3].ToString().ToUpper() + "'";
            DataRow[] resultadoProyecto = dtProyectos.Select(filtro);

            if (resultadoProyecto.Length == 0)
            {
                auxproyecto += dtFacturas.Rows[r][3].ToString().ToUpper();
            }
            else
            {
                projectcode = resultadoProyecto[0]["Projectcode"].ToString().Trim();
            }

            if (auxproyecto != "")
            {
                auxproyecto = auxproyecto.TrimEnd(' ');
                auxproyecto = auxproyecto.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El proyecto " + dtFacturas.Rows[r][3].ToString().ToUpper().Trim() + " no se encuentra dentro de los proyectos activos.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El proyecto " + dtFacturas.Rows[r][3].ToString().ToUpper().Trim() + " no se encuentra dentro de los proyectos activos. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar numero factura
            numfactura = dtFacturas.Rows[r][4].ToString();

            //Verifica si el numero de factura existe en otro registro
            for (int f = 0; f < dtFacturas.Rows.Count; f++)
            {
               if (f!=r) 
               {
                    if (dtFacturas.Rows[f][4].ToString()==numfactura)
                    {
                        auxnumfactura += (f+1).ToString() + "-" + dtFacturas.Rows[f][6].ToString();

                    }
               }
            }


            if (auxnumfactura != "")
            {
                auxnumfactura = auxnumfactura.TrimEnd(' ');
                auxnumfactura = auxnumfactura.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La factura  " + dtFacturas.Rows[r][4].ToString().ToUpper() + " - " + dtFacturas.Rows[r][6].ToString().ToUpper()  + " esta repetido en el registro " + auxnumfactura  + " .");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: La factura  " + dtFacturas.Rows[r][4].ToString().ToUpper() + " - " + dtFacturas.Rows[r][6].ToString().ToUpper()  + " esta repetido en el registro " + auxnumfactura + ".|| ";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion


            // Establecer la cultura a español (España)
            CultureInfo spanishCulture = new CultureInfo("es-ES");
            #region verificar fecha factura



            if (dtFacturas.Rows[r][5].ToString() != "")
            {

                if (dtFacturas.Rows[r][5].ToString().Length >= 10)
                {
                    fechafactura = dtFacturas.Rows[r][5].ToString().Substring(0, 10);

                    DateTime parsed;

                    if (DateTime.TryParseExact(fechafactura.Trim(), "dd/MM/yyyy",
                                    spanishCulture, DateTimeStyles.None, out parsed))
                    {
                        fechafactura = parsed.ToShortDateString();
                    }
                    else
                    {
                        ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La fecha de factura " + dtFacturas.Rows[r][5].ToString().ToUpper().Trim() + " no tiene el formato correcto.");
                        errorExcel += "    - [ERROR Excel: Fila " + r + "]: La fecha de factura " + dtFacturas.Rows[r][5].ToString().ToUpper().Trim() + " no tiene el formato correcto. ||";
                        numErrorExcel++;
                        numerrorfactura++;
                    }
                }
                else
                {
                    ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La fecha de factura " + dtFacturas.Rows[r][5].ToString().ToUpper().Trim() + " no tiene el formato correcto.");
                    errorExcel += "    - [ERROR Excel: Fila " + r + "]: La fecha de factura " + dtFacturas.Rows[r][5].ToString().ToUpper().Trim() + " no tiene el formato correcto. ||";
                    numErrorExcel++;
                    numerrorfactura++;
                }
            }
            else
            {
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La fecha de factura " + dtFacturas.Rows[r][5].ToString().ToUpper().Trim() + " no tiene el formato correcto.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: La fecha de factura " + dtFacturas.Rows[r][5].ToString().ToUpper().Trim() + " no tiene el formato correcto. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion


            concepto = dtFacturas.Rows[r][6].ToString();

            #region verificar moneda

            filtro = "AlphaCode = '" + dtFacturas.Rows[r][7].ToString().ToUpper() + "'";
            DataRow[] resultadoMoneda = dtMonedas.Select(filtro);

            if (resultadoMoneda.Length == 0)
            {
                auxmoneda += dtFacturas.Rows[r][7].ToString().ToUpper();
            }
            else
            {
                moneda = resultadoMoneda[0]["AlphaCode"].ToString();
            }

            if (auxmoneda != "")
            {
                auxmoneda = auxmoneda.TrimEnd(' ');
                auxmoneda = auxmoneda.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La moneda " + dtFacturas.Rows[r][7].ToString().ToUpper().Trim() + " no se encuentra dentro de las monedas activos.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: La moneda " + dtFacturas.Rows[r][7].ToString().ToUpper().Trim() + " no se encuentra dentro de las monedas activos. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar formapago


            if (dtFacturas.Rows[r][8].ToString().ToUpper().Trim() != formapago.ToUpper().Replace(" ","").Trim())
            {
                auxformapago += dtFacturas.Rows[r][8].ToString().ToUpper();
            }


            if (auxformapago != "")
            {
                auxformapago = auxformapago.TrimEnd(' ');
                auxformapago = auxformapago.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La forma de pago " + dtFacturas.Rows[r][8].ToString().ToUpper().Trim() + " no es igual a la forma del pago del proveedor.(" + formapago + ")");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: La forma de pago " + dtFacturas.Rows[r][8].ToString().ToUpper().Trim() + " no es igual a la forma del pago del proveedor.(" + formapago + ")||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion


            dtCuentaOrigen = Payments.GetCuentaOrigen(empresaid);


            #region verificar cuentaorigen

            filtro = "CCC = '" + dtFacturas.Rows[r][9].ToString().ToUpper().ToString() + "'";
            DataRow[] resultadoCuentaOrigen = dtCuentaOrigen.Select(filtro);

            if (resultadoCuentaOrigen.Length == 0)
            {
                auxcuentaorigen += dtFacturas.Rows[r][9].ToString().ToUpper();
            }
            else
            {
                cuentaorigen = resultadoCuentaOrigen[0]["IBAN"].ToString();
   
            }

            if (auxcuentaorigen != "")
            {
                auxcuentaorigen = auxcuentaorigen.TrimEnd(' ');
                auxcuentaorigen = auxcuentaorigen.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La cuenta origen " + dtFacturas.Rows[r][9].ToString().ToUpper().Trim() + " no se encuentra dentro de esa empresa.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: La cuenta origen " + dtFacturas.Rows[r][9].ToString().ToUpper().Trim() + " no se encuentra dentro de esa empresa. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            importebruto = dtFacturas.Rows[r][10].ToString();
            tipoiva = dtFacturas.Rows[r][11].ToString();
            tasa = dtFacturas.Rows[r][12].ToString();
            IRPF = dtFacturas.Rows[r][13].ToString();


            #region verificar fecha fechaorevistapago

            if (dtFacturas.Rows[r][14].ToString() != "")
            {

                if (dtFacturas.Rows[r][14].ToString().Length >= 10)
                {
                    fechaorevistapago = dtFacturas.Rows[r][14].ToString().Substring(0, 10);

                    DateTime parsed2;

                    if (DateTime.TryParseExact(fechaorevistapago, "dd/MM/yyyy",
                                    spanishCulture, DateTimeStyles.None, out parsed2))
                    {
                        fechaorevistapago = parsed2.ToShortDateString();
                    }
                    else
                    {
                        ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La fecha prevista de pago " + dtFacturas.Rows[r][14].ToString().ToUpper().Trim() + " no tiene el formato correcto.");
                        errorExcel += "    - [ERROR Excel: Fila " + r + "]:  La fecha prevista de pago " + dtFacturas.Rows[r][14].ToString().ToUpper().Trim() + " no tiene el formato correcto. ||";
                        numErrorExcel++;
                        numerrorfactura++;
                    }
                }
                else
                {
                    ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La fecha prevista de pago " + dtFacturas.Rows[r][14].ToString().ToUpper().Trim() + " no tiene el formato correcto.");
                    errorExcel += "    - [ERROR Excel: Fila " + r + "]:  La fecha prevista de pago " + dtFacturas.Rows[r][14].ToString().ToUpper().Trim() + " no tiene el formato correcto. ||";
                    numErrorExcel++;
                    numerrorfactura++;
                }
            }
            else
            {
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: La fecha prevista de pago " + dtFacturas.Rows[r][14].ToString().ToUpper().Trim() + " no tiene el formato correcto.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]:  La fecha prevista de pago " + dtFacturas.Rows[r][14].ToString().ToUpper().Trim() + " no tiene el formato correcto. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar tipo gasto


            filtro = "Nombre = '" + dtFacturas.Rows[r][15].ToString().ToUpper() + "'";
            DataRow[] resultadotipogasto = dtTipoGasto.Select(filtro);

            if (resultadotipogasto.Length == 0)
            {
                auxtipogasto += dtFacturas.Rows[r][15].ToString().ToUpper();
            }
            else
            {
                idtipogasto = resultadotipogasto[0]["ID"].ToString().Trim();
                tipogasto = dtFacturas.Rows[r][15].ToString();
            }

            if (auxtipogasto != "")
            {
                auxtipogasto = auxtipogasto.TrimEnd(' ');
                auxtipogasto = auxtipogasto.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El tipo de gasto " + dtFacturas.Rows[r][15].ToString().ToUpper().Trim() + " no esta permitido.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El tipo de gasto " + dtFacturas.Rows[r][15].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar directoindirecto

            dtDirectoIndirecto = Payments.GetGastoDirectoIndirecto(Convert.ToInt32(idtipogasto));

            filtro = "Nombre = '" + dtFacturas.Rows[r][16].ToString().ToUpper() + "'";
            DataRow[] resultadodirectoindirecto = dtDirectoIndirecto.Select(filtro);

            if (resultadodirectoindirecto.Length == 0)
            {
                auxdirectoindirecto += dtFacturas.Rows[r][16].ToString().ToUpper();
            }
            else
            {
                iddirectoindirecto = resultadodirectoindirecto[0]["ID"].ToString().Trim();
                directoindirecto = dtFacturas.Rows[r][16].ToString();
            }

            if (auxdirectoindirecto != "")
            {
                auxdirectoindirecto = auxdirectoindirecto.TrimEnd(' ');
                auxdirectoindirecto = auxdirectoindirecto.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El directo/indirecto " + dtFacturas.Rows[r][16].ToString().ToUpper().Trim() + " no esta permitido.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El directo/indirecto " + dtFacturas.Rows[r][16].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar grupo

            dtGrupo = Payments.GetGrupoGasto(Convert.ToInt32(iddirectoindirecto));

            filtro = "NOMBRE = '" + dtFacturas.Rows[r][17].ToString().ToUpper() + "'";
            DataRow[] resultadogrupo = dtGrupo.Select(filtro);

            if (resultadogrupo.Length == 0)
            {
                auxgrupo += dtFacturas.Rows[r][17].ToString().ToUpper();
            }
            else
            {
                idgrupo = resultadogrupo[0]["ID"].ToString().Trim();
                clavegrupo = resultadogrupo[0]["CLAVE"].ToString().Trim();
                grupo = dtFacturas.Rows[r][17].ToString();
            }

            if (auxgrupo != "")
            {
                auxgrupo = auxgrupo.TrimEnd(' ');
                auxgrupo = auxgrupo.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El grupo " + dtFacturas.Rows[r][17].ToString().ToUpper().Trim() + " no esta permitido.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El grupo " + dtFacturas.Rows[r][17].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar SUBTIPO

            dtSubtipo = Payments.GetSubtipoGasto(Convert.ToInt32(idgrupo));

            filtro = "NOMBRE = '" + dtFacturas.Rows[r][18].ToString().ToUpper() + "'";
            DataRow[] resultadosubtipo = dtSubtipo.Select(filtro);

            if (resultadosubtipo.Length == 0)
            {
                auxsubtipogasto += dtFacturas.Rows[r][18].ToString().ToUpper();
            }
            else
            {
                idsubtipogasto = resultadosubtipo[0]["ID"].ToString().Trim();
                clavesubtipogasto = resultadosubtipo[0]["CLAVE"].ToString().Trim();
                subtipogasto = dtFacturas.Rows[r][18].ToString();

                dtCCCta = Payments.GetCuentaGasto(Convert.ToInt32(idsubtipogasto));
                if(dtCCCta.Rows.Count>0)
                {
                    cccta= dtCCCta.Rows[0]["NumCuenta"].ToString();
                }
           

            }

            if (auxsubtipogasto != "")
            {
                auxsubtipogasto = auxsubtipogasto.TrimEnd(' ');
                auxsubtipogasto = auxsubtipogasto.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El subtipo " + dtFacturas.Rows[r][18].ToString().ToUpper().Trim() + " no esta permitido.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El subtipo " + dtFacturas.Rows[r][18].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar Depto

            dtDepartamento = Departamentos.GetDepartamentosActive(empresaid);

            filtro = "CompoundName = '" + dtFacturas.Rows[r][19].ToString().ToUpper() + "'";
            DataRow[] resultadodepto = dtDepartamento.Select(filtro);

            if (resultadodepto.Length == 0)
            {
                auxdepto += dtFacturas.Rows[r][18].ToString().ToUpper();
            }
            else
            {
                iddepto = resultadodepto[0]["IdDepartamento"].ToString().Trim();
                departamento = resultadodepto[0]["CompoundName"].ToString().Trim();
            }

            if (auxdepto != "")
            {
                auxdepto = auxdepto.TrimEnd(' ');
                auxdepto = auxdepto.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El departamento " + dtFacturas.Rows[r][19].ToString().ToUpper().Trim() + " no esta permitido.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El departamento " + dtFacturas.Rows[r][19].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            #region verificar estado


            if (dtFacturas.Rows[r][20].ToString().Trim() != "CREADA" &&
                dtFacturas.Rows[r][20].ToString().Trim() != "RECHAZADA" &&
                dtFacturas.Rows[r][20].ToString().Trim() != "PENDIENTE")
            {
                auxestado += dtFacturas.Rows[r][20].ToString().ToUpper();
            }
            else
            {
                estado = dtFacturas.Rows[r][20].ToString();
            }

            if (auxestado != "")
            {
                auxestado = auxestado.TrimEnd(' ');
                auxestado = auxestado.TrimEnd(',');
                ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: El estado " + dtFacturas.Rows[r][20].ToString().ToUpper().Trim() + " no esta permitido.");
                errorExcel += "    - [ERROR Excel: Fila " + r + "]: El estado " + dtFacturas.Rows[r][20].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                numErrorExcel++;
                numerrorfactura++;
            }

            #endregion

            if (tipogasto == "INVERSION" &&
               directoindirecto == "DIRECTO" &&
               grupo == "Directo - Inversiones" &&
               subtipogasto == "APPS")
            {
                compraapps = true;
            }

            if (compraapps)
            {

                #region verificar entidad financiadora


                if (dtFacturas.Rows[r][22].ToString().ToUpper() != "")
                {

                    filtro = "ID = '" + dtFacturas.Rows[r][22].ToString().ToUpper() + "'";
                    DataRow[] resultadentidadfinanciadora = dtEntidadFinanciadora.Select(filtro);

                    if (resultadentidadfinanciadora.Length == 0)
                    {
                        auxentidadfinanciada += dtFacturas.Rows[r][22].ToString().ToUpper();
                    }
                    else
                    {
                        entidadfinanciadora = dtFacturas.Rows[r][22].ToString();
                    }

                    if (auxentidadfinanciada != "")
                    {
                        auxentidadfinanciada = auxentidadfinanciada.TrimEnd(' ');
                        auxentidadfinanciada = auxentidadfinanciada.TrimEnd(',');
                        ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: Entidad financiadora " + dtFacturas.Rows[r][22].ToString().ToUpper().Trim() + " no esta permitido.");
                        errorExcel += "    - [ERROR Excel: Fila " + r + "]:  Entidad financiadora " + dtFacturas.Rows[r][22].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                        numErrorExcel++;
                        numerrorfactura++;
                    }
                }
                else
                {
                    auxentidadfinanciada = auxentidadfinanciada.TrimEnd(' ');
                    auxentidadfinanciada = auxentidadfinanciada.TrimEnd(',');
                    ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: Entidad financiadora " + dtFacturas.Rows[r][22].ToString().ToUpper().Trim() + " no esta permitido.");
                    errorExcel += "    - [ERROR Excel: Fila " + r + "]:  Entidad financiadora " + dtFacturas.Rows[r][22].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                    numErrorExcel++;
                    numerrorfactura++;
                }



                #endregion

                #region verificar codigo contrato

                if (dtFacturas.Rows[r][23].ToString().ToUpper() != "")
                {
                    filtro = "ContractCode = '" + dtFacturas.Rows[r][23].ToString().ToUpper() + "'";
                    DataRow[] resultadocodigocontrato = dtCodigoContrato.Select(filtro);

                    if (resultadocodigocontrato.Length == 0)
                    {
                        auxcodigocontrato += dtFacturas.Rows[r][23].ToString().ToUpper();
                    }
                    else
                    {
                        codigocontrato = dtFacturas.Rows[r][23].ToString();
                    }

                    if (auxcodigocontrato != "")
                    {
                        auxcodigocontrato = auxcodigocontrato.TrimEnd(' ');
                        auxcodigocontrato = auxcodigocontrato.TrimEnd(',');
                        ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: Codigo contrato " + dtFacturas.Rows[r][23].ToString().ToUpper().Trim() + " no esta permitido.");
                        errorExcel += "    - [ERROR Excel: Fila " + r + "]:  Codigo contrato " + dtFacturas.Rows[r][23].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                        numErrorExcel++;
                    }
                }
                else
                {
                    auxcodigocontrato = auxcodigocontrato.TrimEnd(' ');
                    auxcodigocontrato = auxcodigocontrato.TrimEnd(',');
                    ASPxListBoxWarnings.Items.Add("[ERROR Excel: Fila " + r + "]: Codigo contrato " + dtFacturas.Rows[r][23].ToString().ToUpper().Trim() + " no esta permitido.");
                    errorExcel += "    - [ERROR Excel: Fila " + r + "]:  Codigo contrato " + dtFacturas.Rows[r][23].ToString().ToUpper().Trim() + " no esta permitdo. ||";
                    numErrorExcel++;
                    numerrorfactura++;
                }

                #endregion

            }

            if (numerrorfactura == 0)
            {
                numokfactura++;
                ASPxListBoxWarnings.Items.Add("[OK Excel: Fila " + r + "]: Factura " + dtFacturas.Rows[r][4].ToString().ToUpper().Trim() + " validada correctamente.");

                row = DaTa.NewRow();
                row["tipoproveedor"] = tipoproveedor;
                row["empresaid"] = empresaid;
                row["empresa"] = empresa;
                row["providerid"] = providerid;
                row["providername"] = providername;
                row["CIF"] = CIF;
                row["CCC"] = CCC;
                row["SWIFT"] = SWIFT;
                row["IBAN"] = IBAN;
                row["ABA"] = ABA;
                row["AccountNumber"] = AccountNumber;
                row["BankName"] = BankName;


                row["projectcode"] = projectcode;
                row["numfactura"] = numfactura;
                row["fechafactura"] = fechafactura;
                row["concepto"] = concepto;
                row["moneda"] = moneda;
                row["formapago"] = formapago;
                row["cuentaorigen"] = cuentaorigen;
                row["importebruto"] = importebruto;
                row["tipoiva"] = tipoiva;
                row["IRPF"] = IRPF;
                row["fechaorevistapago"] = fechaorevistapago;
                row["tasa"] = tasa;
                row["idtipogasto"] = idtipogasto;
                row["tipogasto"] = tipogasto;
                row["iddirectoindirecto"] = iddirectoindirecto;
                row["directoindirecto"] = directoindirecto;
                row["idgrupo"] = idgrupo;
                row["grupo"] = grupo;
                row["clavegrupo"] = clavegrupo;
                row["idsubtipogasto"] = idsubtipogasto;
                row["subtipogasto"] = subtipogasto;
                row["clavesubtipogasto"] = clavesubtipogasto;
                row["cccta"] = cccta;
                row["iddepto"] = iddepto;
                row["departamento"] = departamento;
                row["estado"] = estado;
                row["nombredocumento"] = nombredocumento;
                row["entidadfinanciadora"] = entidadfinanciadora;
                row["codigocontrato"] = codigocontrato;
                row["bankaccountid"] = bankaccountid;
                row["noSujeto"] = noSujetoData1;

                DaTa.Rows.Add(row);


            }

        }


        if (numerrorfactura > 0)
        {
            ASPxLabelMsg.Text = "Revise el listado de errores.";
            ASPxLabelMsg.Visible = true;
            ASPxButtonProcesar.Visible = false ;
        }
        else
        {
            ASPxGridViewFacturas.DataSource = DaTa;
            ASPxGridViewFacturas.DataBind();
            ASPxButtonProcesar.Visible = true;
        }



    }


    protected void ASPxButtonProcesar_Click(object sender, EventArgs e)
    {


        #region variables 
        string tipoproveedor = "";
        string empresaid = "";
        string empresa = "";
        string providerid = "";
        string providername = "";
        string CIF = "";
        string CCC = "";
        string SWIFT = "";
        string IBAN = "";
        string ABA = "";
        string AccountNumber = "";
        string BankName = "";
        string projectcode = "";
        string numfactura = "";
        string fechafactura = "";
        string concepto = "";
        string moneda = "";
        string formapago = "";
        string cuentaorigen = "";
        string importebruto = "";
        string tipoiva = "";
        string tasa = "";
        string IRPF = "";
        string fechaorevistapago = "";
        string idtipogasto = "";
        string tipogasto = "";
        string iddirectoindirecto = "";
        string directoindirecto = "";

        string idgrupo = "";
        string grupo = "";
        string clavegrupo = "";

        string idsubtipogasto = "";
        string subtipogasto = "";
        string clavesubtipogasto = "";
        string cccta = "";

        string iddepto = "";
        string departamento = "";
        string estado = "";
        string nombredocumento = "";
        string extensiondocumento = "";
        string entidadfinanciadora = "";
        string codigocontrato = "";
        string bankaccountid = "";

        //constantes
        string observaciones = "Carga masiva desde Excel";
        string postedby = HttpContext.Current.Profile.UserName.ToString();
        decimal cuota = 0;
        string carencia = "0";
        string tipologiapago = "TOTAL";
        string motivo = "-1";
        string fechaconstituacion = DateTime.Today.ToString();
        string fechafin = DateTime.Today.ToString();

        string noSujeto = "";

        string ccpais = "01";
        string ccoficina = "0";
        string cccliente = "000";

        ASPxListBoxWarnings.Items.Clear();
        ASPxListBoxWarnings.Items.Add("Procesamiento carga masiva de facturas");
        ASPxListBoxWarnings.Items.Add("------------------------------------------------");

        #endregion

        for (int r = 0; r < DaTa.Rows.Count; r++)
        {

            tipoproveedor = DaTa.Rows[r]["tipoproveedor"].ToString();
            empresaid = DaTa.Rows[r]["empresaid"].ToString();
            empresa = DaTa.Rows[r]["empresa"].ToString();
            providerid =  DaTa.Rows[r]["providerid"].ToString();
            providername = DaTa.Rows[r]["providername"].ToString();
            CIF = DaTa.Rows[r]["CIF"].ToString();
            CCC = DaTa.Rows[r]["CCC"].ToString();
            SWIFT = DaTa.Rows[r]["SWIFT"].ToString();
            IBAN = DaTa.Rows[r]["IBAN"].ToString();
            ABA = DaTa.Rows[r]["ABA"].ToString();
            AccountNumber = DaTa.Rows[r]["AccountNumber"].ToString();
            BankName = DaTa.Rows[r]["BankName"].ToString();



            projectcode = DaTa.Rows[r]["projectcode"].ToString();
            numfactura = DaTa.Rows[r]["numfactura"].ToString();
            fechafactura = DaTa.Rows[r]["fechafactura"].ToString();
            concepto = DaTa.Rows[r]["concepto"].ToString();
            moneda = DaTa.Rows[r]["moneda"].ToString();
            formapago = DaTa.Rows[r]["formapago"].ToString();
            cuentaorigen = DaTa.Rows[r]["cuentaorigen"].ToString();
            importebruto = DaTa.Rows[r]["importebruto"].ToString();
            tipoiva = DaTa.Rows[r]["tipoiva"].ToString();
            tasa = DaTa.Rows[r]["tasa"].ToString();
            IRPF = DaTa.Rows[r]["IRPF"].ToString();
            fechaorevistapago = DaTa.Rows[r]["fechaorevistapago"].ToString();
            idtipogasto = DaTa.Rows[r]["idtipogasto"].ToString();
            tipogasto = DaTa.Rows[r]["tipogasto"].ToString();
            iddirectoindirecto = DaTa.Rows[r]["iddirectoindirecto"].ToString();
            directoindirecto = DaTa.Rows[r]["directoindirecto"].ToString();

            idgrupo = DaTa.Rows[r]["idgrupo"].ToString();
            grupo = DaTa.Rows[r]["grupo"].ToString();
            clavegrupo = DaTa.Rows[r]["clavegrupo"].ToString();

            idsubtipogasto = DaTa.Rows[r]["idsubtipogasto"].ToString();
            subtipogasto = DaTa.Rows[r]["subtipogasto"].ToString();
            clavesubtipogasto = DaTa.Rows[r]["clavesubtipogasto"].ToString();
            cccta= DaTa.Rows[r]["cccta"].ToString();

            noSujeto = DaTa.Rows[r]["noSujeto"].ToString();

            iddepto = DaTa.Rows[r]["iddepto"].ToString();
            departamento = DaTa.Rows[r]["departamento"].ToString();
            estado = DaTa.Rows[r]["estado"].ToString();
            nombredocumento = DaTa.Rows[r]["nombredocumento"].ToString();

            if (string.IsNullOrEmpty(IRPF))
            {
                IRPF = "0";
            }
            string directoryPath = HttpContext.Current.Server.MapPath("~/App_Data/tmp/subidamasiva/");

            string[] posiblesExtensiones = { ".docx", ".doc", ".pdf", ".xlsx" };
            string extension = Path.GetExtension(nombredocumento);
            if (string.IsNullOrEmpty(extension))
            {
                foreach (string ext in posiblesExtensiones)
                {
                    // Combina el nombre del archivo con cada extensión posible
                    string tempFilePath = Path.Combine(directoryPath, nombredocumento + ext);

                    // Comprueba si el archivo con esa extensión existe
                    if (File.Exists(tempFilePath))
                    {
                        extension = ext;
                        nombredocumento += extension;
                        break;
                    }
                }
            }

            // Construir la ruta completa del archivo
            string filename = Path.Combine(directoryPath, nombredocumento);
            FileStream fs = new FileStream(filename,
                                   FileMode.Open,
                                   FileAccess.Read);
            BinaryReader br = new BinaryReader(fs);
            long numBytes = new FileInfo(filename).Length;
            fichero_contenido = br.ReadBytes((int)numBytes);
 


            extensiondocumento = Path.GetExtension(nombredocumento);
            entidadfinanciadora = DaTa.Rows[r]["entidadfinanciadora"].ToString();
            codigocontrato = DaTa.Rows[r]["codigocontrato"].ToString();
            bankaccountid = DaTa.Rows[r]["bankaccountid"].ToString();

            //calculos bruto
            decimal brutosinva = 0;
            decimal brutoiva = 0;
            decimal importeiva = 0;
            string iva = "0";
            decimal netoiva = 0;

            decimal brutoirpf = 0;
            decimal importeirpf = 0;
            string irpf = IRPF;
            decimal netoirpf = 0;

            decimal totalbruto = 0;
            decimal totaltasas = 0;
            decimal total = 0;
            decimal noSujetoBase = 0;




            //AQUI ESTA EL CALCULO DE IRPF----------------------------------------------------------

            if(tipoiva=="IVA")
            {
                tasa = tasa.Replace(".", ",");
                iva = tasa;
                brutosinva = 0;
                brutoiva = decimal.Parse(importebruto);
               decimal prueba  = decimal.Parse(iva) ;
                importeiva = decimal.Parse(iva) * brutoiva;
                netoiva = brutoiva + importeiva;
                brutoirpf = decimal.Parse(importebruto); ;
                importeirpf = brutoirpf * decimal.Parse(IRPF);
                irpf = IRPF;
                netoirpf = importeirpf;
                totalbruto = brutoiva;
                totaltasas = importeiva + importeirpf;
                if (noSujeto == "")
                {
                    noSujetoBase = 0;
                }
                else {
                    noSujetoBase = decimal.Parse(noSujeto);
                }

                
                total = totalbruto + importeiva - importeirpf + noSujetoBase;

            }
            else
            {
                
                //iva = irpf; // por qué iva es igual a irpf?
                brutosinva = 0;
                brutoiva = decimal.Parse(importebruto); 
                importeiva = 0;
                importeiva = decimal.Parse(iva) * brutoiva;
                netoiva = brutoiva + importeiva;
                brutoirpf = decimal.Parse(importebruto);
                importeirpf = decimal.Parse(irpf) * brutoirpf;
                //irpf = irpf;
                netoirpf = decimal.Parse(iva) * brutoiva;
                totalbruto = brutoiva;
                totaltasas = importeiva+importeirpf;
                if (noSujeto == "")
                {
                    noSujetoBase = 0;
                }
                else
                {
                    noSujetoBase = decimal.Parse(noSujeto);
                }
                total = totalbruto + importeiva - importeirpf + noSujetoBase;
            }


            #region UPDATE BBDD PAGO 
            int PaymentId = -1;

            try
            {
                PaymentId = Payments.UpdatePagoGenerico(
                    0,
                    tipoproveedor,
                    empresaid,
                    providerid,
                    projectcode,
                    CIF,
                    providername,
                    CCC,
                    CCC,
                    SWIFT,
                    IBAN,
                    ABA,
                    AccountNumber,
                    BankName,
                    DateTime.Parse(fechaorevistapago),
                    numfactura,
                    fechafactura,
                    concepto,
                    formapago.ToUpper(),
                    tipologiapago.ToUpper(),
                    "",
                    moneda,
                    tipogasto,
                    clavesubtipogasto,
                    directoindirecto.ToUpper(),
                    brutosinva,
                    brutoiva,
                    importeiva,
                    decimal.Parse(iva),
                    netoiva,
                    brutoirpf,
                    importeirpf,
                    decimal.Parse(irpf),
                    netoirpf,
                    totalbruto,
                    totaltasas,
                    total,
                    0,
                    total,
                    ccpais,
                    ccoficina,
                    iddepto,
                    clavegrupo,
                    cccta,
                    cccliente,
                    estado,
                    observaciones,
                    fichero_contenido, //nombredocumento
                    extensiondocumento,
                    postedby,
                    motivo,
                    tipoiva,
                    fechaconstituacion,
                    fechafin,
                    cuota,
                    carencia,
                    null,
                    null,
                    null,
                    null,
                    int.Parse(bankaccountid),
                    0);

           

                if (PaymentId!=-1)
                {
                    

                    if (Payments.InsertPartialPayment(PaymentId, fechaorevistapago,"", total, postedby) != -1)
                    {
                        DateTime fechapagoreal=DateTime.Now;
                        Payments.InsertJustificantePago(PaymentId, nombredocumento, extensiondocumento, fichero_contenido, fechapagoreal, total, moneda, postedby);

                    }

                    ASPxListBoxWarnings.Items.Add("Factura Nº:" + numfactura + " con cencepto " + concepto + " de fecha " + fechafactura + " creada correctamente. PaymentId: " + PaymentId );
                }
            }
            catch (Exception ex)
            {
                ASPxListBoxWarnings.Items.Add("**ERROR : al crear la factura Nº:" + numfactura + " con cencepto" + concepto + " de fecha " + fechafactura + "Error: " + ex.Message);
                throw new Exception("Exception" + ex.Message);
            }


            #endregion

        }



        DaTa.Clear();
    }


    protected string getUser()
    {
        string tokenFinal = "";
        try
        {
            WebRequest theRequest = WebRequest.Create("http://api.linksnowpeak.com:5001/login");
            theRequest.Method = "POST";

            var byteArray = Encoding.UTF8.GetBytes("sotec.emanage:weewoo2020");

            theRequest.Headers.Add("user-info", Convert.ToBase64String(byteArray));
            Stream requestStream = theRequest.GetRequestStream();

            requestStream.Close();

            HttpWebResponse response = (HttpWebResponse)theRequest.GetResponse();
            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            Token_Inversiones token = serializer.Deserialize<Token_Inversiones>(responseString);

            tokenFinal = token.x_api_key;

            return tokenFinal;
        }
        catch (Exception e)
        {
            throw new Exception("Exception throw at getUser: " + e.Message);
        }
    }

    protected List<Payments_Contratos> getContracts(string token)
    {
        List<Payments_Contratos> contracts = new List<Payments_Contratos>();
        try
        {
            WebRequest theRequest = WebRequest.Create("http://api.linksnowpeak.com:5001/contract_deals/contract_deals_list/");
            theRequest.Method = "GET";

            theRequest.Headers.Add("x-api-key", token);

            using (WebResponse response = theRequest.GetResponse())
            {
                using (Stream stream = response.GetResponseStream())
                {
                    var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                    var settings = new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    };
                    var list = contracts.Concat<Payments_Contratos>(contracts);
                    contracts = contracts.Concat<Payments_Contratos>(JsonConvert.DeserializeObject<List<Payments_Contratos>>(responseString, settings)).ToList();
                }
            }

            return contracts;
        }
        catch (Exception e)
        {
            throw new Exception("Exception throw at get contracts: " + e.Message);
        }
    }
}

