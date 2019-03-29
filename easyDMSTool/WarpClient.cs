using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Text;
using System.Xml;
using System.Net;
using System.IO;

namespace easyDMSTool
{
    class WarpClient
    {        
        private const int PRIORITY = 6;
        private const String APPLICATION = "EasySender";
        private const String EXPORT_SCHEME_FINANCE_WORKFLOW = "/Finance-Workflow";
        private const String EXPORT_SCHEME_CS_WORKFLOW = "/CustomerService-Workflow";
        private const String QA_SUFFIX = "-QA";
        private const String SCAN_USER = "mailbox export";
        private const String URL = "/ccdocupload";

        private static String server;
        private static String port;


        private String docType;
        private String countryCode;
        private String info;
        private String userName;


        public bool  sendDocument2EDMS(String docType, String cc, String information, String exportScheme, String currentUser, String fileName, long fileSize, String fileBase64) {

            String docString = xml2string(buildXML(docType, cc, information, exportScheme, currentUser, fileName, fileSize, fileBase64));
            return postXMLData("http://deis044:4420/ccdocupload", docString);            
        }
        

        private static XDocument buildXML(String docType, String cc, String info, string scheme, String userName, String attachmentName, long attachmentSize, string fileBinary)
        {

            XElement document = new XElement("ccDocument", new XAttribute("Id", ""), new XAttribute("StackId", ""),
                new XElement("Options",
                    new XElement("Priority", PRIORITY),
                    new XElement("UserName", userName),
                    new XElement("Application", APPLICATION),
                    new XElement("ExportScheme", scheme)),
                new XElement("Content",
                    new XElement("Index",
                        new XElement("Field", new XAttribute("Name", "Doc_Type"), docType),
                        new XElement("Field", new XAttribute("Name", "Country_Code"), cc),
                        new XElement("Field", new XAttribute("Name", "Scan_User"), SCAN_USER),
                        new XElement("Field", new XAttribute("Name", "Information"), info)
                ),
                    new XElement("Blobs", new XElement("Blob",
                         new XElement("Name", attachmentName),
                         new XElement("MimeType", "application/pdf"),
                         new XElement("Extension", ".pdf"),
                         new XElement("Size", attachmentSize),
                         new XElement("Data", new XCData(fileBinary))
                    ))

                ));

            XDocument xdoc = new XDocument(
                 new XDeclaration("1.0", "utf-8", "yes"),
                 document);

            return xdoc;
        }

        private string xml2string(XDocument doc)
        {
            return doc.Declaration.ToString() + Environment.NewLine + doc.ToString();
        }

        private bool postXMLData(string destinationUrl, string requestXml)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(destinationUrl);
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(requestXml);
            request.ContentType = "text/xml";
            request.ContentLength = bytes.Length;
            request.Method = "PUT";

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();

            HttpWebResponse response;
            try
            {
                response = (HttpWebResponse)request.GetResponse();
                string responseStr = "";
                if (response.StatusCode == HttpStatusCode.Created)
                {
                    System.IO.Stream responseStream = response.GetResponseStream();
                    responseStr = new StreamReader(responseStream).ReadToEnd();
                }
                return true;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                return false;
            }
        }
    }
}
