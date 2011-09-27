using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SwDocumentMgr;

namespace GetSolidWorksFileVersion
{
    class Program
    {
        static void Main(string[] args)
        {
            string docPath;
            bool quietMode;
            switch (args.Length)
            {
                case 1:
                    quietMode = false;
                    docPath = args[0];
                    break;
                case 2:
                    if (args[0] != "/q")
                    {
                        quietMode = false;
                        inputError(quietMode);
                        return;
                    }
                    quietMode = true;
                    docPath = args[1];
                    break;
                default:
                    quietMode = false;
                    inputError(quietMode);
                    return;
            }

            //Get Document Type
            SwDmDocumentType docType = setDocType(docPath);



            SwDMClassFactory dmClassFact = new SwDMClassFactory();
            SwDMApplication dmDocManager = dmClassFact.GetApplication("C3021E19A05D3A4B15A8D40AD7B0474E7FD073393806CD0E") as SwDMApplication;
            SwDmDocumentOpenError OpenError;
            SwDMDocument dmDoc = dmDocManager.GetDocument(docPath, docType, true, out OpenError) as SwDMDocument;

            int fileVersion;

            if (dmDoc != null)
            {
                try
                {
                    fileVersion = dmDoc.GetVersion();
                    if (fileVersion == 0)
                    {
                        Console.WriteLine(docPath + "\t  Build " + fileVersion + ".  This probably isn't a SolidWorks file or it is severely damaged.");
                        return;
                    }
                    Console.WriteLine(docPath + "\t SolidWorks Build " + fileVersion);
                }
                catch (Exception e)
                {
                    Console.WriteLine(docPath + "\t" + "Could not get file version. " + e.Message + e.StackTrace);
                    return;
                }
                dmDoc.CloseDoc();
            }
            else
            {
                switch (OpenError)
                {
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorFail:
                        Console.WriteLine(docPath + "\tFile failed to open; reasons could be related to permissions, the file is in use, or the file is corrupted.");
                        inputError(quietMode);
                        break;
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorFileNotFound:
                        Console.WriteLine(docPath + "\tFile not found");
                        inputError(quietMode);
                        break;
                    case SwDmDocumentOpenError.swDmDocumentOpenErrorNonSW:
                        Console.Write(docPath + "\tNon-SolidWorks file was opened");
                        inputError(quietMode);
                        break;
                    default:
                        Console.WriteLine(docPath + "\tAn unknown errror occurred.  Something is wrong with the code of \"GetSolidWorksFileVersion\"");
                        inputError(quietMode);
                        break;
                }
            }

        }




        static SwDmDocumentType setDocType(string docPath)
        {
            string fileExtension = System.IO.Path.GetExtension(docPath);

            //Notice no break statement is needed because I used return to get out of the switch statement.
            switch (fileExtension.ToUpper())
            {
                case ".SLDASM":
                case ".ASM":
                    return SwDmDocumentType.swDmDocumentAssembly;
                case ".SLDDRW":
                case ".DRW":
                    return SwDmDocumentType.swDmDocumentDrawing;
                case ".SLDPRT":
                case ".PRT":
                    return SwDmDocumentType.swDmDocumentPart;
                default:
                    return SwDmDocumentType.swDmDocumentUnknown;
            }

        }





        static void inputError(bool quietMode)
        {
            if (quietMode)
                return;

            Console.WriteLine(@"

Returns the Version of SolidWorks Version.
Syntax 
    [option] [SolidWorksFilePath]

Output
    ""SolidWorksFilePath""  BuildNumber

Some SolidWorks Build Numbers:
SolidWorks 97+      629
SolidWorks 98	    817
SolidWorks 98+      1008
SolidWorks 99	    1137
SolidWorks 2000	    1500
SolidWorks 2001	    1750
SolidWorks 2001+    1950
SolidWorks 2003	    2200
SolidWorks 2004	    2500
SolidWorks 2005	    2800
SolidWorks 2006	    3100
SolidWorks 2007	    3400
SolidWorks 2008	    3800
SolidWorks 2009	    4100
SolidWorks 2010	    4400
SolidWorks 2011	    4700

Version 2011-Sept-26 22:42
Written and Maintained by Jason Nicholson
http://github.com/jasonnicholson/GetSolidWorksFileVersion");
        }

    }
}
