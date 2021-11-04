using Access_data.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace ALS.Controllers
{
    [RoutePrefix("api/report")]
    public class ReportController : ApiController
    {
        ReportService report = new ReportService();
        [HttpGet]
        [Route("report01/{PACKAGENO}")]
        public IHttpActionResult report01(string PACKAGENO)
        {
            try
            {
                var result = report.Report01(PACKAGENO);
                HttpResponseMessage response;
                response = Request.CreateResponse(HttpStatusCode.OK);
                //octet - stream
                MediaTypeHeaderValue mediaType = new MediaTypeHeaderValue("application/pdf");
                response.Content = new StreamContent(result);
                response.Content.Headers.ContentType = mediaType;
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                response.Content.Headers.ContentDisposition.FileName = "report01.pdf";
                return ResponseMessage(response);
            }
            catch (Exception ex)
            {
                return Json(new { is_error = true, error_message = string.Format("{0}", ex.Message) });
            }
        }

        [HttpGet]
        [Route("report02/{BANO}")]
        public IHttpActionResult report02(string BANO)
        {
            try
            {
                var result = report.Report02(BANO);
                HttpResponseMessage response;
                response = Request.CreateResponse(HttpStatusCode.OK);
                //octet - stream
                MediaTypeHeaderValue mediaType = new MediaTypeHeaderValue("application/pdf");
                response.Content = new StreamContent(result);
                response.Content.Headers.ContentType = mediaType;
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                response.Content.Headers.ContentDisposition.FileName = "report01.pdf";
                return ResponseMessage(response);
            }
            catch (Exception ex)
            {
                return Json(new { is_error = true, error_message = string.Format("{0}", ex.Message) });
            }
        }
    }
}
