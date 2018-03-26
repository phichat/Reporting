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
                var receiveNo = Request.QueryString["ReceiveNo"];
                var reportType = Request.QueryString["ReportType"];
                switch (rptName)
                {
                    case "PartsReceive":
                        PartsReceive(receiveNo, reportType);
                        break;

                    case "StockAvailable":
                        StockAvailable(receiveNo, reportType);
                        break;

                    case "StockMaterial":
                        string shop = Request.QueryString["Shop"];
                        DateTime dateFrom = DateTime.Parse(Request.QueryString["DateFrom"]);
                        DateTime dateTo = DateTime.Parse(Request.QueryString["DateTo"]);

                        StockMaterial(shop, dateFrom.Date, dateTo, reportType);
                        break;

                    case "PartsMovement":
                        var partNo = Request.QueryString["PartNo"];
                        PartsMovement(receiveNo, partNo, reportType);
                        break;

                    case "CarsMovement":
                        CarsMovement(receiveNo, reportType);
                        break;
                }
            }
        }

        private void PartsReceive(string receiveNo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            try
            {
                conn.Open();
                string sqlQry = "EXEC dbo.sp_rptPartsReceive '" + receiveNo + "'";

                var dt = new DataTable();
                var da = new SqlDataAdapter(sqlQry, conn);
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./PartsReceive.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@ReceiveNo", receiveNo);
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

        private void PartsMovement(string receiveNo, string partNo, string reportType)
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
                cmd.Parameters.AddWithValue("@ReceiveNo", receiveNo);
                cmd.Parameters.AddWithValue("@PartNo", partNo);
                cmd.Connection = conn;

                da.SelectCommand = cmd;
                da.Fill(dt);

                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./PartsMovement.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@ReceiveNo", receiveNo);
                rptDoc.SetParameterValue("@PartNo", partNo);
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
                    ? ExportFormatType.Excel
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

        private void StockMaterial(string shop, DateTime dateFrom, DateTime dateTo, string reportType)
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
                cmd.Parameters.AddWithValue("@Shop", shop);
                cmd.Parameters.AddWithValue("@DateFrom", dateFrom.ToString("yyyy-MM-dd"));
                cmd.Parameters.AddWithValue("@DateTo", dateTo.ToString("yyyy-MM-dd"));
                cmd.Connection = conn;

                da.SelectCommand = cmd;
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./StockMaterial.rpt"));
                rptDoc.Refresh();
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@Shop", shop);
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

        private void CarsMovement(string receiveNo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            rptDoc = new ReportDocument();
            try
            {
                conn.Open();
                string sqlQry = "EXEC dbo.sp_rptCarsMovement '" + receiveNo + "'";

                var dt = new DataTable();
                var da = new SqlDataAdapter(sqlQry, conn);
                da.Fill(dt);

                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./CarsMovement.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("@ReceiveNo", receiveNo);
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Cars-Movement");

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