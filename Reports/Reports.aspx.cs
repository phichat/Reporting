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

                string sqlQry = "SELECT DISTINCT";
                sqlQry += " DateToProduction";
                sqlQry += " , CustomEntryNo";
                sqlQry += " , InvoiceNo";
                sqlQry += " , ReceiveNo";
                sqlQry += " , PartNo";
                sqlQry += " , PartDescription";
                sqlQry += " , Qty";
                sqlQry += " , QPV";
                sqlQry += " , Amount";
                sqlQry += " , UM";
                sqlQry += " FROM dbo.PartReceive";
                sqlQry += " where ReceiveNo = '" + receiveNo + "'";

                var dt = new DataTable();
                var da = new SqlDataAdapter(sqlQry, conn);
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                var file = (reportType == "pdf"
                    ? Server.MapPath("./PartsReceivePDF.rpt")
                    : Server.MapPath("./PartsReceiveExcel.rpt"));

                rptDoc.Load(file);
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("ReceiveNo", receiveNo);
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
            try
            {
                conn.Open();
                string sqlQry = " SELECT DISTINCT";
                sqlQry += "  D.ReceiveDate";
                sqlQry += "  , Consignment";
                sqlQry += "  , H.ReceiveNo";
                sqlQry += "  , PartNo";
                sqlQry += "  , PartDescription";
                sqlQry += "  , Qty";
                sqlQry += "  , QPV";
                sqlQry += "  , H.CommissionNo";
                sqlQry += "  , H.[Date] AS DateFG";
                sqlQry += "  , H.[Date] AS DateBuyOff";
                sqlQry += "  , H.VDONo";
                sqlQry += "  , CASE WHEN VDONo IS NULL THEN 1 ELSE 0 END Amount";
                sqlQry += " FROM dbo.ReceiveReference AS H";
                sqlQry += " LEFT JOIN dbo.PartReceive AS D ON D.ReceiveNo = H.ReceiveNo";
                sqlQry += " WHERE D.ReceiveNo = '" + receiveNo + "' AND D.PartNo = '" + partNo + "'";
                sqlQry += " ORDER BY D.ReceiveDate, D.PartNo, CommissionNo";

                var dt = new DataTable();
                var da = new SqlDataAdapter(sqlQry, conn);
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./PartsMovement.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("ReceiveNo", receiveNo);
                rptDoc.SetParameterValue("PartNo", partNo);
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Parts-Movement");

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

        private void StockAvailable(string receiveNo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            try
            {
                conn.Open();
                string sqlQry = "SELECT DISTINCT";
                sqlQry += " DateToProduction";
                sqlQry += " , CustomEntryNo";
                sqlQry += " , InvoiceNo";
                sqlQry += " , ReceiveNo";
                sqlQry += " , PartNo";
                sqlQry += " , PartDescription";
                sqlQry += " , Qty";
                sqlQry += " , QPV";
                sqlQry += " , Amount";
                sqlQry += " , UM";
                sqlQry += " FROM dbo.PartReceive";
                sqlQry += " where ReceiveNo = '" + receiveNo + "'";

                var dt = new DataTable();
                var da = new SqlDataAdapter(sqlQry, conn);
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./SotckAvailable.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("ReceiveNo", receiveNo);
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Stock-Available");

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

        private void StockMaterial(string shop, DateTime dateFrom, DateTime dateTo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            try
            {
                conn.Open();
                string sqlQry = " SELECT";
                sqlQry += " ReceiveDate";
                sqlQry += " , ReceiveNo";
                sqlQry += " , CustomEntryNo";
                sqlQry += " , InvoiceNo";
                sqlQry += " , CommissionFrom";
                sqlQry += " , CommissionTo";
                sqlQry += " , Model";
                sqlQry += " , Shop";
                sqlQry += " , sum(Qty) Qty";
                sqlQry += " FROM dbo.PartReceive";
                sqlQry += " where Shop = '" + shop + "'";
                sqlQry += " and cast(ReceiveDate as date) between '" + dateFrom.ToString("yyyy-MM-dd") + "'";
                sqlQry += " and '" + dateTo.ToString("yyyy-MM-dd") + "'";
                sqlQry += " group by ReceiveDate";
                sqlQry += " , ReceiveNo";
                sqlQry += " , CustomEntryNo";
                sqlQry += " , InvoiceNo";
                sqlQry += " , CommissionFrom";
                sqlQry += " , CommissionTo";
                sqlQry += " , Model";
                sqlQry += " , Shop";

                var dt = new DataTable();
                var da = new SqlDataAdapter(sqlQry, conn);
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./StockMaterial.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("Shop", shop);
                rptDoc.SetParameterValue("dateFrom", dateFrom.ToString("yyyy-MM-dd"));
                rptDoc.SetParameterValue("dateTo", dateTo.ToString("yyyy-MM-dd"));
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Stock-Material");

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

        private void CarsMovement(string receiveNo, string reportType)
        {
            conn = new SqlConnection(connTaap);
            try
            {
                conn.Open();
                string sqlQry = "SELECT DISTINCT";
                sqlQry += "   D.[Date] AS DateFG";
                sqlQry += " , D.CommissionNo";
                sqlQry += " , H.Model";
                sqlQry += " , H.PackingMonth";
                sqlQry += " , Consignment";
                sqlQry += " , H.QPV";
                sqlQry += " , D.[Date] AS DateBuyOff";
                sqlQry += " , Amount";

                sqlQry += " FROM db_Taap.dbo.PartReceive AS H";
                sqlQry += " LEFT JOIN dbo.ReceiveReference AS D";
                sqlQry += " ON D.ReceiveNo = H.ReceiveNo";
                sqlQry += " WHERE H.ReceiveNo = '" + receiveNo + "' and D.[Status] = 1";

                var dt = new DataTable();
                var da = new SqlDataAdapter(sqlQry, conn);
                da.Fill(dt);

                rptDoc = new ReportDocument();
                var rptReponse = (reportType == "excel"
                    ? ExportFormatType.Excel
                    : ExportFormatType.PortableDocFormat);

                rptDoc.Load(Server.MapPath("./CarsMovement.rpt"));
                rptDoc.SetDataSource(dt);
                rptDoc.SetParameterValue("ReceiveNo", receiveNo);
                rptDoc.ExportToHttpResponse(rptReponse, Response, true, "Report-Cars-Movement");

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

    }
}