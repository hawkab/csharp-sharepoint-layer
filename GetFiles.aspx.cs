using System;
using System.Text;
using System.Net;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebPartPages;


namespace GetCalendar.Layouts.GetCalendar
{
    public partial class Calendar : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string[] parameters = new string[] { Request.QueryString.Get("oid"), Request.QueryString.Get("fromdate"), Request.QueryString.Get("todate") };

            int oid = 0;
            DateTime fromDate, toDate;
            int.TryParse(parameters[0], out oid);
            
            fromDate = (parameters[1] == null) ? DateTime.Now.AddDays(-14) : Convert.ToDateTime(parameters[1]);
            toDate = (parameters[2] == null) ? DateTime.Now.AddDays(160) : Convert.ToDateTime(parameters[2]);

            if (parameters[0] == null)
            {
                Response.StatusCode = 411; Response.End();
            }
            else
            {
                if (oid != 0)
                {
                    string url = "http://remoteserver/files/service.svc/getfile?fromdate=" + fromDate.ToString("yyyy-MM-dd") + "&todate=" + toDate.ToString("yyyy-MM-dd") + "&receivertype=1&lectureroid=" + oid;
                    WebClient client = new WebClient();

                    string strContentType = "application/octet-stream";
                    byte[] obj = client.DownloadData(url);
                    string fileName = "";
                    if (!String.IsNullOrEmpty(client.ResponseHeaders["Content-Disposition"]))
                    {
                        int i = client.ResponseHeaders["Content-Disposition"].IndexOf("filename*=");
                        if (i > 0)
                        {
                            fileName = client.ResponseHeaders["Content-Disposition"].Substring(i + 17).Replace("\"", "");
                            Response.ClearContent();
                            Response.ClearHeaders();
                            Response.AppendHeader("Content-Disposition", "attachment; filename= " + fileName);
                            Response.ContentType = strContentType;

                            if (Response.IsClientConnected)
                                Response.BinaryWrite(obj);
                            Response.Flush();
                            Response.Close();
                        }
                    }
                    else { Response.StatusCode = 204; Response.End(); }
                }
                else { Response.StatusCode = 400; Response.End(); }
            }
        }
    }
}
