using System;
using System.Diagnostics;
using System.IO;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;

namespace OpenOfficeCli
{
    class Program
    {

        public static void Main(string[] args)
        {
            Console.WriteLine("Ready to go?");
            Console.WriteLine("1)Docx");
            Console.WriteLine("2)Doc");
            Console.WriteLine("3)Xlsx");
            Console.WriteLine("4)Xls");
            Console.WriteLine("5)Pptx");
            Console.WriteLine("6)Ppt");
            var x = Console.ReadLine();

            string fromFile = null;
            string toFile1 = null;
            string toFile2 = null;
            switch (x)
            {
                case "1":
                    fromFile = @"C:\test\demo2.docx";

                    toFile1 = @"C:\test\Export\WordxPDF.pdf";
                    toFile2 = @"C:\test\Export\WordxODT.odt";

                    ConvertToPdf(fromFile, toFile1);
                    ConvertToPdf(fromFile, toFile2);
                    break;
                case "3":
                    fromFile = @"C:\test\demo.xlsx";

                    toFile1 = @"C:\test\Export\ExcelxPDF.pdf";
                    toFile2 = @"C:\test\Export\ExcelxODS.ods";

                    ConvertToPdf(fromFile, toFile1);
                    ConvertToPdf(fromFile, toFile2);
                    break;
                case "5":
                    fromFile = @"C:\test\demo1.pptx";

                    toFile1 = @"C:\test\Export\PPTxPDF.pdf";
                    toFile2 = @"C:\test\Export\PPTxODS.odp";

                    ConvertToPdf(fromFile, toFile1);
                    ConvertToPdf(fromFile, toFile2);
                    break;
                case "6":
                    fromFile = @"C:\test\demo1.ppt";

                    toFile1 = @"C:\test\Export\PPTPDF.pdf";
                    toFile2 = @"C:\test\Export\PPTODS.odp";

                    ConvertToPdf(fromFile, toFile1);
                    ConvertToPdf(fromFile, toFile2);
                    break;
                default:
                    Console.WriteLine("cancel");
                    break;

            }

            //string fromFile = @"file://C:/test/importDoc.docx";
            //string toFile = @"file:///C:/test/exportPDF.docx";
            //string fromFile = @"C:\test\demo1.docx";
            //string toFile1 = @"C:\test\exportPDF.pdf";
            //string toFile2 = @"C:\test\exportPDF.odt";

            //ConvertToPdf(fromFile, toFile1);
            //ConvertToPdf(fromFile, toFile2);

        }

        public static void ConvertToPdf(string inputFile, string outputFile)
        {
            if (ConvertExtensionToFilterType(Path.GetExtension(inputFile)) == null)
                throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + inputFile);

            //StartOpenOffice();

            //Get a ComponentContext
            var xLocalContext = Bootstrap.bootstrap();
            //Get MultiServiceFactory
            XMultiServiceFactory xRemoteFactory = (XMultiServiceFactory)xLocalContext.getServiceManager();
            //Get a CompontLoader
            XComponentLoader aLoader = (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
            //Load the sourcefile

            XComponent xComponent = null;
            try
            {
                xComponent = initDocument(aLoader, PathConverter(inputFile), "_blank");
                //Wait for loading
                while (xComponent == null)
                {
                    System.Threading.Thread.Sleep(1000);
                }

                // save/export the document
                saveDocument(xComponent, inputFile, PathConverter(outputFile));

            }
            catch { throw; }
            finally { xComponent.dispose(); }
        }

        //private static void StartOpenOffice()
        //{
        //    Process[] ps = Process.GetProcessesByName("soffice.exe");
        //    if (ps != null)
        //    {
        //        if (ps.Length > 0)
        //            return;
        //        else
        //        {
        //            Process p = new Process();
        //            p.StartInfo.Arguments = "-headless -nofirststartwizard";
        //            p.StartInfo.FileName = "soffice.exe";
        //            p.StartInfo.CreateNoWindow = true;
        //            bool result = p.Start();
        //            if (result == false)
        //                throw new InvalidProgramException("OpenOffice failed to start.");
        //        }
        //    }
        //    else
        //    {
        //        throw new InvalidProgramException("OpenOffice not found.  Is OpenOffice installed?");
        //    }
        //}


        private static XComponent initDocument(XComponentLoader aLoader, string file, string target)
        {
            PropertyValue[] openProps = new PropertyValue[1];
            openProps[0] = new PropertyValue();
            openProps[0].Name = "Hidden";
            openProps[0].Value = new uno.Any(true);


            XComponent xComponent = aLoader.loadComponentFromURL(
            file, target, 0,
            openProps);

            return xComponent;
        }


        private static void saveDocument(XComponent xComponent, string sourceFile, string destinationFile)
        {
            var sourceFileType = Path.GetExtension(sourceFile);
            var destinationFileType = Path.GetExtension(destinationFile);

            PropertyValue[] propertyValues = new PropertyValue[2];

            propertyValues[0] = new PropertyValue { Name = "FilterName", Value = new uno.Any(ConvertExtensionToFilterType(destinationFileType, sourceFileType)) };
            propertyValues[1] = new PropertyValue { Name = "Overwrite", Value = new uno.Any(true) };

            ((XStorable)xComponent).storeToURL(destinationFile, propertyValues);


            //PropertyValue[] propVals = new PropertyValue[0];
            //XComponent oDoc = oDesk.loadComponentFromURL(url, "_blank", 0, propVals);
        }


        private static string PathConverter(string file)
        {
            if (file == null || file.Length == 0)
                throw new NullReferenceException("Null or empty path passed to OpenOffice");

            return String.Format("file:///{0}", file.Replace(@"\", "/"));

        }

        public static string ConvertExtensionToFilterType(string extension, string inputFileType = null)
        {


            switch (extension)
            {
                case ".odt":
                    return "writer8";
                case ".pdf":
                    {
                        switch (inputFileType)
                        {
                            case ".doc":
                            case ".docx":
                                return "writer_pdf_Export";
                            case ".xlsx":
                            case ".xls":
                            case ".xlsb":
                                return "calc_pdf_Export";
                            case ".ppt":
                            case ".pptx":
                                return "impress_pdf_Export";
                            default:
                                return "writer_pdf_Export";
                        }
                    }
                case ".doc":
                case ".docx":
                case ".txt":
                case ".rtf":
                case ".html":
                case ".htm":
                case ".xml":
                case ".wps":
                case ".wpd":
                    return "writer_pdf_Export";
                case ".ods":
                    return "calc8";
                case ".xlsx":
                case ".xls":
                case ".xlsb":
                    return "calc_pdf_Export";
                case ".odp":
                    return "impress8";
                case ".ppt":
                case ".pptx":
                    return "impress_pdf_Export";
                default: return null;
            }
        }
    }
}
