using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using ExportFile.Helpers;
using ExportFile.Models;

namespace ExportFile.Controllers
{
    public class ExportController : ApiController
    {
        readonly List<UserModel> _users = new List<UserModel>
        {
            new UserModel { FirstName = "Jane", LastName = "Doe", DateOfBirth = DateTime.Now.AddYears(-70) },
            new UserModel { FirstName = "John", LastName = "Doe", DateOfBirth = DateTime.Now.AddYears(-65) },
            new UserModel { FirstName = "Jane", LastName = "Doe", DateOfBirth = DateTime.Now.AddYears(-60) },
            new UserModel { FirstName = "John", LastName = "Doe", DateOfBirth = DateTime.Now.AddYears(-55) },
            new UserModel { FirstName = "Jane", LastName = "Doe", DateOfBirth = DateTime.Now.AddYears(-50) },
            new UserModel { FirstName = "John", LastName = "Doe", DateOfBirth = DateTime.Now.AddYears(-45) },
            new UserModel { FirstName = "Jane", LastName = "Doe", DateOfBirth = DateTime.Now.AddYears(-40) },
            new UserModel { FirstName = "John", LastName = "Doe", DateOfBirth = DateTime.Now.AddYears(-35) }
        };

        [HttpGet, Route("api/ExportToCsv")]
        public IHttpActionResult ExportToCsv()
        {
            // Convert the list of users to a DataTable => easy to manipulate column names
            var dtUsers = _users.ToDataTable();
            // Convert the DataTable to an array of bytes
            var content = dtUsers.ToCsvByteArray();
            // Send the array of bytes (file) to FE
            var responseMsg = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(content)
            };
            responseMsg.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "users.csv"
            };
            responseMsg.Content.Headers.ContentType = new MediaTypeHeaderValue("text/csv");
            IHttpActionResult response = ResponseMessage((responseMsg));
            return response;
        }

        [HttpGet, Route("api/ExportToXls")]
        public IHttpActionResult ExportToXls()
        {
            // Convert the list of users to a DataTable => easy to manipulate column names
            var dtUsers = _users.ToDataTable();
            // Convert the DataTable to an array of bytes
            var contentStream = dtUsers.ToExcelStream();

            var responseMsg = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(contentStream)
            };
            responseMsg.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "users.xls"
            };
            responseMsg.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");
            IHttpActionResult response = ResponseMessage((responseMsg));
            return response;
        }

        [HttpGet, Route("api/ExportToXlsx")]
        public IHttpActionResult ExportToXlsx()
        {
            // Convert the list of users to a DataTable => easy to manipulate column names
            var dtUsers = _users.ToDataTable();
            // Convert the DataTable to an array of bytes
            var content = dtUsers.ToXlsxByteArray();

            var responseMsg = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(content)
            };
            responseMsg.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "users.xlsx"
            };
            responseMsg.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            IHttpActionResult response = ResponseMessage((responseMsg));
            return response;
        }
    }
}
