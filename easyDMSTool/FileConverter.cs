using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Compression;
using System.Web.Services.Protocols;
using System.Data.SqlClient;
using easyDMSTool.Doc4SOAP;
using Base64Tools;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using easyDMSTool.ToolBox;
using easyDMSTool.Converter;
using System.Reflection;
using System.Threading;
using easyDMSTool.Properties;
using System.ServiceModel;
using Microsoft.Office.Interop.Word;


namespace easyDMSTool
{
    class FileConverter
    {
        private ToolBox.ToolboxService t = null;
        private ToolBox.toolboxOptions o = null;
        private Converter.ConverterService cc = null;
        private Base64Encoder base64Encoder = null;
        private string workDir;
        private string instanceURL;
        private bool filesOK = true;
        private Boolean isSentToEasy = true;
        private Boolean isConverted = true;
        private int count = 0;
        public bool Convert(string directory, string outputFile, string fileType, string country, string docType, string url, string countryCode, string emailSender)
        {
            this.workDir = !directory.EndsWith(@"\") ? (directory + @"\") : directory;
            this.instanceURL = "http://" + url + ":11001";
            this.filesOK = this.processFiles();
            if (!this.filesOK)
            {
                return false;
            }
            this.convertFiles();
            return isConverted;
        }

        private void convertFiles()
        {
            foreach (string file in Directory.GetFiles(this.workDir, "*"))
            {
                if ((Path.GetExtension(file) != ".pdf") && (Path.GetExtension(file) != ".exe"))
                {
                    string extension = Path.GetExtension(file);
                    byte[] fileContent = this.ReadByteArray(file);
                    try
                    {
                        if ((extension.ToLower() == ".xls") | (extension.ToLower() == ".xlsx"))
                        {
                            this.convertExcel(file);
                        }
                        else if (extension == ".txt")
                        {
                            File.Move(file, Path.ChangeExtension(file, ".docx"));
                        }
                        else if (extension == ".rtf")
                        {
                            this.saveRFTasPDF(file);
                        }
                        else if ((extension != ".eml") && (extension != ".msg") && ((extension!="") | (extension != null)))
                        {
                            this.cc = new ConverterService();
                            byte[] p = this.cc.convertDocumentSimple(extension, ".pdf", "pdf.archive=true&pdf.embedFonts=true&pdfa.level=2a&reportContentProblems=true", fileContent);
                            string fileName = this.workDir + Path.GetFileNameWithoutExtension(file) + ".pdf";
                            this.WriteByteArray(p, fileName);
                        }
                    }
                    catch (SoapException exception)
                    {
                        MessageBox.Show("SOAP conversion service error!\nSoapException in convertFile method: " + exception.Message + "\n in file: " + Path.GetFileName(file));
                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show("Conversion service error!\nException in convertFile method: " + exception.Message + "\n in file: " + Path.GetFileName(file));
                    }
                }
            
            }
        }


        private void convertExcel(string fileName)
        {
            string str = Path.GetFileNameWithoutExtension(fileName) + ".pdf";
            Application application = (Application)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
            Workbook workbook = null;
            object missing = Type.Missing;
            string filename = this.workDir + str;
            XlFixedFormatType xlTypePDF = XlFixedFormatType.xlTypePDF;
            XlFixedFormatQuality xlQualityMinimum = XlFixedFormatQuality.xlQualityMinimum;
            bool openAfterPublish = false;
            bool includeDocProperties = true;
            bool ignorePrintAreas = true;
            object from = Type.Missing;
            object to = Type.Missing;
            try
            {
                workbook = application.Workbooks.Open(fileName, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                if (workbook != null)
                {
                    workbook.ExportAsFixedFormat(xlTypePDF, filename, xlQualityMinimum, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, missing);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Excel conversion failed for file " + fileName + ". \n Error caught: " + exception.Message);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false, missing, missing);
                    workbook = null;
                }
                if (application != null)
                {
                    application.Quit();
                    application = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        public string mergeFiles()
        {
            string baseFile = "";
            string[] files;
            files = Directory.GetFiles(this.workDir, "Message_Body.pdf");
                if (files.Length <= 0)
                {
                    baseFile = Directory.GetFiles(this.workDir, "*.pdf").Last<string>();
                }
                else
                {
                    baseFile = (Directory.GetFiles(this.workDir, "Message_Body.pdf").Length != 1) ? files.First<string>() : Directory.GetFiles(this.workDir, "Message_Body.pdf")[0];
                    this.mergeFiles(baseFile, files);
                }
                this.mergeFiles(baseFile, Directory.GetFiles(this.workDir, "*.pdf"));
                return baseFile;
          
        }    

        private void mergeFiles(string baseFile, string[] rest)
        {
            foreach (string str in rest)
            {
                if (str != baseFile)
                {
                    this.mergeFiles(baseFile, str);
                    File.Delete(str);
                }
            }
        }

        private void mergeFiles(string file1, string file2)      
        {
            this.t = new ToolboxService();
            this.o = new toolboxOptions();
            string p = file1;
            this.base64Encoder = new Base64Encoder(this.ReadByteArray(file2));
            this.o.serviceOptions = "pdf.pdfOperation=7";
            this.o.toolboxOptions1 = $"pdf.merge.pdfFile={new string(this.base64Encoder.GetEncoded())}";
            byte[] fileContent = this.ReadByteArray(p);
            byte[] buffer2 = this.t.processDocument(this.o, fileContent);
            this.WriteByteArray(buffer2, p);
        }


        private bool moveOutput(string outputFile, string fileType, string country, string docType, string countryCode, string sourceFile, string emailSender)
        {
            string session = null;
            if (!File.Exists(sourceFile))
            {
                sourceFile = this.workDir + "Message_Body.pdf";
                MessageBox.Show("PLEASE CHECK DOC IN EASYDMS!\n{0}", emailSender);
            }
            bool flag = false;
            DOCUMENTSPortTypeClient client = new DOCUMENTSPortTypeClient("DOCUMENTS", this.instanceURL);
            FieldData[] fields = new FieldData[] { new FieldData(), new FieldData(), new FieldData(), new FieldData(), new FieldData() };
            fields[0].name = "Doc_Type";
            fields[0].value = docType;
            fields[1].name = "Country_Code";
            fields[1].value = countryCode;
            fields[2].name = "Status";
            fields[2].value = "new";
            fields[3].name = "Scan_User";
            fields[3].value = "mailbox auto export";
            fields[4].name = "Information";
            fields[4].value = emailSender;
            DocUploadData[] addDocs = new DocUploadData[] { new DocUploadData() };
            addDocs[0].name = "exported_email.pdf";
            addDocs[0].register = "attachments";
            while (true)
            {
                while (true)
                {
                    try
                    {
                        char[] trimChars = new char[] { ' ', '\n' };
                        char[] chArray2 = new char[] { ' ', '\n' };
                        char[] chArray3 = new char[] { ' ', '\n' };
                        client.trustedLogin(Settings.Default.userIDDefault.TrimEnd(trimChars), "72765", Settings.Default.userPasswordDefault.TrimEnd(chArray2), ribbonEasyDMS.getUserID().TrimEnd(chArray3), "", "en", out session);
                        addDocs[0].data = File.ReadAllBytes(sourceFile);
                        client.createFile(ref session, fileType, fields, addDocs);
                        flag = true;
                        this.isSentToEasy = true;
                    }
                    //TOD: handle login issue
                    catch (FaultException fex)
                    {
                        MessageBox.Show("Cought some exception \n" + fex.Message);
                        Thread.Sleep(500);
                        this.isSentToEasy = false;
                        this.count++;
                    }
                    catch (SoapException)
                    {
                        Thread.Sleep(500);
                        this.isSentToEasy = false;
                        this.count++;
                    }
                    catch (DirectoryNotFoundException exception)
                    {
                        MessageBox.Show("DirectoryNotFoundException Error creating document in Easy DMS! \n" + exception.Message);
                        flag = true;
                        this.isSentToEasy = false;
                    }
                    catch (Exception exception2)
                    {
                        MessageBox.Show("Exception Error creating document in Easy DMS! \n" + exception2.Message);
                        this.count++;
                        if (this.count > 5)
                        {
                            flag = true;
                        }
                        this.isSentToEasy = false;
                    }
                    finally
                    {
                        if ((session != null) && flag)
                        {
                            client.logout(ref session);
                            client.Close();
                            if (Directory.Exists(this.workDir))
                            {
                                Directory.Delete(this.workDir, true);
                            }
                        }
                        if (this.count > 8)
                        {
                            this.isSentToEasy = false;
                        }
                    }
                    break;
                }
                if (flag && (this.count <= 8))
                {
                    return this.isSentToEasy;
                }
            }
        }

        private bool processFiles()
        {
            bool flag = true;
            foreach (string str in Directory.GetFiles(this.workDir, "*"))
            {
                if (Path.GetExtension(str).ToLower() == ".zip")
                {
                    try
                    {
                        ZipStorer storer = ZipStorer.Open(str, FileAccess.Read);
                        foreach (ZipStorer.ZipFileEntry entry in storer.ReadCentralDir())
                        {
                            storer.ExtractFile(entry, this.workDir + entry.FilenameInZip);
                        }
                        storer.Close();
                        File.Delete(str);
                    }
                    catch (Exception)
                    {
                        flag = false;
                    }
                }
            }
            foreach (string str2 in Directory.GetFiles(this.workDir, "*"))
            {
                if ((Path.GetExtension(str2).ToLower() == ".tif") || (Path.GetExtension(str2).ToLower() == ".tiff"))
                {
                    if (!GraphicsManipulation.ConvertTiffToJpeg(str2))
                    {
                        flag = false;
                    }
                    File.Delete(str2);
                }
                else if (Path.GetExtension(str2).ToLower() == ".bmp")
                {
                    if (GraphicsManipulation.ConvertBmpToJpeg(str2))
                    {
                        flag = false;
                    }
                    File.Delete(str2);
                }
            }
            foreach (string str3 in Directory.GetFiles(this.workDir, "*"))
            {
                if (((Path.GetExtension(str3).ToLower() == ".jpg") || (Path.GetExtension(str3).ToLower() == ".jpeg")) && !GraphicsManipulation.ShrinkJPEG(str3, (long)50))
                {
                    flag = false;
                }
            }
            return flag;
        }
          

        private void saveRFTasPDF(string fileName)
        {
            string outputFileName = this.workDir + Path.GetFileNameWithoutExtension(fileName) + ".pdf";
            Microsoft.Office.Interop.Word._Application objWord = null;
            objWord = (Microsoft.Office.Interop.Word.Application)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
            Document document = objWord.Documents.Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            document.ExportAsFixedFormat(outputFileName, WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, false, true, WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false, Missing.Value);
            document.Save();
            document.Close(Missing.Value, Missing.Value, Missing.Value);
        }

        private byte[] ReadByteArray(string p) =>
        File.ReadAllBytes(p);

        private void WriteByteArray(byte[] p, string fileName)
        {
            File.WriteAllBytes(fileName, p);
        }

    }
}