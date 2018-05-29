using CrystalDecisions.CrystalReports.Engine;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CrystalDecisions.Shared;

namespace Reporting.Reports
{
    public partial class Reports : System.Web.UI.Page
    {

        private string connTaap = ConfigurationManager.ConnectionStrings["TaapMssql"].ConnectionString;
        private SqlConnection conn = null;
        private ReportDocument rptDoc;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(Request.QueryString["ReportName"]))
            {
                var rptName = Request.QueryString["ReportName"];
                var reportType = Request.QueryString["ReportType"];
                switch (rptName)
                {
                    case "PartsReceive":
                        string receiveNo = Request.QueryString["ReceiveNo"];
                        DateTime dateFrom = DateTime.Parse(Request.QueryString["DateFrom"]);
                        DateTime dateTo = DateTime.Parse(Request.QueryString["DateTo"]);
                        PartsReceive(receiveNo, dateFrom, dateTo, reportType);
                        break;

                    case "StockAvailable":
                        receiveNo = Request.QueryString["ReceiveNo"];
                        StockAvailable(receiveNo, reportType);
                        break;

                    case "StockMaterial":
                        string packingMonth = Request.QueryString["PackingMonth"];
                        string model = Request.QueryString["Model"];
                        dateFrom = DateTime.Parse(Request.QueryString["DateFrom"]);
                        dateTo = DateTime.Parse(Request.QueryString["DateTo"]);

                        StockMaterial(packingMonth, model, dateFrom, dateTo, reportType);
                        break;

                    case "PartsMovement":
                        var partNo = Request.QueryString["PartNo"];
                        model = Request.QueryString["Model"];
                        dateFrom = DateTime.Parse(Request.QueryString["DateFrom"]);
                        dateTo = DateTime.Parse(Request.QueryString["DateTo"]);
                        PartsMovement(model, partNo, dateFrom, dateTo, reportType);
                        break;

                    case "CarsMovement":
                        model = Request.QueryString["Model"];
                        packingMonth = Request.QueryString["PackingMonth"];
                        dateFrom = DateTime.Parse(Request.QueryString["DateFrom"]);
                        dateTo = DateTime.Parse(Request.QueryString["DateTo"]);
                        CarsMovement(model, packingMonth, dateFrom, dateTo, reportType);
                        break;
                }
            }
        }

        private void PartsReceive(string receiveNo, DateTime dateFrom, DateTime dateTo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            var cmd = new SqlCommand();
            var dt = new DataTable();
            var da = new SqlDataAdapter();
            rptDoc = new ReportDocument();

            try
            {
                conn.Open();

                cmd.CommandText = "dbo.sp_rptPartsReceive";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ReceiveNo", receiveNo);
                cmd.Parameters.AddWithValue("@Sdate", dateFrom.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@Edate", dateTo.ToString("yyyy-MM-dd"));
                cmd.Connection = conn;

                da.SelectCommand = cmd;
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.ExcelRecord
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./PartsReceive.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@ReceiveNo", receiveNo);
                rptDoc.SetParameterValue("@Sdate", dateFrom.ToString("yyyy-MM-dd"));
                rptDoc.SetParameterValue("@Edate", dateTo.ToString("yyyy-MM-dd"));
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Parts-Receive");
            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
            finally
            {
                conn.Close();
            }

        }

        private void PartsMovement(string model, string partNo, DateTime dateFrom, DateTime dateTo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            var cmd = new SqlCommand();
            var dt = new DataTable();
            var da = new SqlDataAdapter();
            rptDoc = new ReportDocument();

            try
            {
                conn.Open();

                cmd.CommandText = "dbo.sp_rptPartsMovement";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Model", model);
                cmd.Parameters.AddWithValue("@PartNo", partNo);
                cmd.Parameters.AddWithValue("@DateFrom", dateFrom.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@DateTo", dateTo.ToString("yyyy-MM-dd"));
                cmd.Connection = conn;

                da.SelectCommand = cmd;
                da.Fill(dt);

                var rptReponse = (reportType == "excel" ? ExportFormatType.ExcelRecord : ExportFormatType.PortableDocFormat);
                var pathFile = reportType == "excel" ? "./PartsMovementExcel.rpt" : "./PartsMovement.rpt";
                rptDoc.Load(Server.MapPath(pathFile));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@Model", model);
                rptDoc.SetParameterValue("@PartNo", partNo);
                rptDoc.SetParameterValue("@DateFrom", dateFrom.ToString("yyyy-MM-dd"));
                rptDoc.SetParameterValue("@DateTo", dateTo.ToString("yyyy-MM-dd"));
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Parts-Movement");

            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
            finally
            {
                conn.Close();
                rptDoc.Close();
            }
        }

        private void StockAvailable(string receiveNo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            rptDoc = new ReportDocument();
            try
            {
                conn.Open();
                string sqlQry = "dbo.sp_rptStockAvailable '" + receiveNo + "'";

                var dt = new DataTable();
                var da = new SqlDataAdapter(sqlQry, conn);
                da.Fill(dt);

                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.ExcelRecord
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./SotckAvailable.rpt"));
                rptDoc.Refresh();
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@ReceiveNo", receiveNo);
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Stock-Available");

            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
            finally
            {
                conn.Close();
                rptDoc.Close();
            }
        }

        private void StockMaterial(string packingMonth, string model, DateTime dateFrom, DateTime dateTo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            var cmd = new SqlCommand();
            var dt = new DataTable();
            var da = new SqlDataAdapter();
            rptDoc = new ReportDocument();
            try
            {
                conn.Open();

                cmd.CommandText = "dbo.sp_rptStockMaterial";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PackingMonth", packingMonth);
                cmd.Parameters.AddWithValue("@Model", model);
                cmd.Parameters.AddWithValue("@DateFrom", dateFrom.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@DateTo", dateTo.ToString("yyyy-MM-dd"));
                cmd.Connection = conn;

                da.SelectCommand = cmd;
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.ExcelRecord
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./StockMaterial.rpt"));
                rptDoc.Refresh();
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@PackingMonth", packingMonth);
                rptDoc.SetParameterValue("@Model", model);
                rptDoc.SetParameterValue("@DateFrom", dateFrom.ToString("yyyy-MM-dd"));
                rptDoc.SetParameterValue("@DateTo", dateTo.ToString("yyyy-MM-dd"));
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Stock-Material");

            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
            finally
            {
                conn.Close();
                rptDoc.Close();
            }
        }

        private void CarsMovement(string model, string packingMonth, DateTime dateFrom, DateTime dateTo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            var cmd = new SqlCommand();
            var dt = new DataTable();
            var da = new SqlDataAdapter();
            rptDoc = new ReportDocument();
            try
            {
                conn.Open();

                cmd.CommandText = "dbo.sp_rptCarsMovement";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PackingMonth", packingMonth);
                cmd.Parameters.AddWithValue("@Model", model);
                cmd.Parameters.AddWithValue("@DateFrom", dateFrom.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@DateTo", dateTo.ToString("yyyy-MM-dd"));
                cmd.Connection = conn;

                da.SelectCommand = cmd;
                da.Fill(dt);

                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.ExcelRecord
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./CarsMovement.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@PackingMonth", packingMonth);
                rptDoc.SetParameterValue("@Model", model);
                rptDoc.SetParameterValue("@DateFrom", dateFrom.ToString("yyyy-MM-dd"));
                rptDoc.SetParameterValue("@DateTo", dateTo.ToString("yyyy-MM-dd"));
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Cars-Finish-Good");

            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
            finally
            {
                conn.Close();
                rptDoc.Close();
            }
        }

    }
}