using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.Exchange.WebServices.Data;

namespace easyDMSTool
{
    class TraceListener : ITraceListener
    {
        #region ITraceListener Members
        public void Trace(string traceType, string traceMessage)
        {
            CreateXMLTextFile(traceType, traceMessage.ToString());
        }
        #endregion

        private void CreateXMLTextFile(string fileName, string traceContent)
        {
            // Create a new XML file for the trace information.
            try
            {
                // If the trace data is valid XML, create an XmlDocument object and save.

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(traceContent);
                //Debug.WriteLine(traceContent);
                System.Diagnostics.Debug.WriteLine(xmlDoc);
                //xmlDoc.Save("C:\\Users\\plubbart\\Desktop\\xmldocument.xml");
            }
            catch
            {
                // If the trace data is not valid XML, save it as a text document.
                System.IO.File.WriteAllText(fileName + ".txt", traceContent);
            }
        }
    }
}
