using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using log4net;
using System.IO;
using System.Text;

namespace PdfService.Controllers
{
    public class ValuesController : ApiController
    {
        protected static readonly ILog log = LogManager.GetLogger(typeof(ValuesController));

        private static readonly Random random = new Random();

        public static string RandomString(int size, bool lowerCase)
        {
            StringBuilder builder = new StringBuilder();

            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(25 * random.NextDouble() + 65));
                builder.Append(ch);
            }

            if (lowerCase)
            {
                return builder.ToString().ToLower();
            }
            return builder.ToString();
        }

        public string Get(string templateUrl, string headerUrl, string footerUrl, int marginLeft, int marginRight, int marginBottom, int headerSpacing)
        {
            
            return ProcessTemplate(templateUrl, headerUrl, footerUrl, marginLeft, marginRight, marginBottom, headerSpacing, "");    
        }

        public string Get(string templateUrl)
        {
            var result = ProcessTemplate(templateUrl, "", "", 15, 15, 15, 10, "");
            return result;
        }

        public string Get(string templateUrl, string password)
        {
            var result = ProcessTemplate(templateUrl, "", "", 20, 20, 20, 20, password);
            return result;
         }

            // GET api/values/5
        public string Get(string templateUrl, string headerUrl, string footerUrl, int marginLeft, int marginRight, int marginBottom, int headerSpacing, string password)
        {
            return ProcessTemplate(templateUrl, headerUrl, footerUrl, marginLeft, marginRight, marginBottom, headerSpacing, password);
        }

        public string ProcessTemplate(string templateUrl, string headerUrl, string footerUrl, int marginLeft, int marginRight, int marginBottom, int headerSpacing, string password)
        {
            var outputFilenameRoot = RandomString(7, false);
            var outputFilename = outputFilenameRoot + ".pdf";
            string outputPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath(@"~/output"), outputFilename);
            var result = HtmlToPdf(templateUrl, outputPath, headerUrl, footerUrl, marginLeft, marginRight, marginBottom, headerSpacing);

            if (result != 0)
                return "wkhtml2pdf returned error: " + result.ToString();

            if (string.IsNullOrEmpty(password))
            {
                return "/output/" + outputFilename;
            }

            var finalOutputFilename = outputFilenameRoot + "p.pdf";
            var finalOutputPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath(@"~/output"), finalOutputFilename);
            result = AddPasswordToPdf(outputPath, password, finalOutputPath);

            File.Delete(outputPath);

            if (result != 0)
                return "qpdf returned error: " + result.ToString();

            return "/output/" + finalOutputFilename;
        }

        public bool Delete(string fileName)
        { 
            if (fileName.Contains('\\') || fileName.Contains('/'))
                return false;
            string outputPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath(@"~/output"), fileName);
            File.Delete(outputPath);
            return true;
        }

        private static int AddPasswordToPdf(string fileName, string password, string outputFilename)
        {
            var p = new System.Diagnostics.Process();
            p.StartInfo.FileName = @"C:\Program Files (x86)\qpdf-8.0.2\bin\qpdf.exe";
            p.StartInfo.WorkingDirectory = @"C:\Program Files (x86)\qpdf-8.0.2\bin";
            p.StartInfo.UseShellExecute = false; // needs to be false in order to redirect output
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.RedirectStandardInput = true;

            string switches = string.Format(" --encrypt {0} {0} 40 -- ", password, password);

            var completeArguments = switches + " " + fileName + " " + outputFilename;
            p.StartInfo.Arguments = completeArguments;
            log.Info("AddPasswordToPdf: " + completeArguments);

            p.StartInfo.UseShellExecute = false; // needs to be false in order to redirect output
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.RedirectStandardInput = true; // redirect all 3, as it should be all 3 or none
            p.Start();

            string output = p.StandardOutput.ReadToEnd();
            p.WaitForExit(60000);
            int returnCode = p.ExitCode;
            p.Close();

            return returnCode;
        }

        /// <summary>
        /// Convert Html page at a given URL to a PDF file using open-source tool wkhtml2pdf
        /// </summary>
        /// <param name="Url"></param>
        /// <param name="outputFilename"></param>
        /// <returns></returns>
        private static int HtmlToPdf(string Url, string outputFilename, string headerUrl, string footerUrl, int marginLeft, int marginRight, int marginBottom, int headerSpacing)
        {
            // assemble destination PDF file name

            // get proj no for header
            //Project project = new Project(int.Parse(outputFilename));

            var p = new System.Diagnostics.Process();
            p.StartInfo.FileName = @"C:\Program Files (x86)\wkhtmltopdf\wkhtmltopdf.exe";
            p.StartInfo.WorkingDirectory = @"C:\Program Files (x86)\wkhtmltopdf";
            //p.StartInfo.WorkingDirectory = @"E:\dev\netrebel\Balirka\Balirka.Web\files";

            // --disable-smart-shrinking

            p.StartInfo.UseShellExecute = false; // needs to be false in order to redirect output
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.RedirectStandardInput = true; // redirect all 3, as it should be all 3 or none
            //-B 0 -L 0 -R 0 -T 0
            string switches = " -s A4  --margin-right " + marginRight + "mm --margin-left " + marginLeft + "mm --margin-bottom " + marginBottom + "mm ";
            
            // --header-spacing " + headerSpacing + "mm";

            switches += !string.IsNullOrWhiteSpace(headerUrl) ? " --header-html " + headerUrl + "" : " ";
            switches += !string.IsNullOrWhiteSpace(footerUrl) ? " --footer-html " + footerUrl + "" : " ";

            //string switches = "--print-media-type ";
            //switches += "--margin-top 4mm --margin-bottom 4mm --margin-right 0mm --margin-left 0mm ";
            //switches += "--page-size A4 ";
            //switches += "--no-background ";
            //switches += "--redirect-delay 100";

            var completeArguments = switches + " \"" + Url + "\" " + outputFilename;
            p.StartInfo.Arguments = completeArguments;
            log.Info("HtmlToPdf: " + completeArguments);

            p.StartInfo.UseShellExecute = false; // needs to be false in order to redirect output
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.RedirectStandardInput = true; // redirect all 3, as it should be all 3 or none


            p.Start();

            // read the output here...
            string output = p.StandardOutput.ReadToEnd();

            // ...then wait n milliseconds for exit (as after exit, it can't read the output)
            p.WaitForExit(60000);

            // read the exit code, close process
            int returnCode = p.ExitCode;
            p.Close();

            return returnCode;
        }
    }
}