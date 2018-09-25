using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Web;

public class Handler : IHttpHandler
{
    public void ProcessRequest(HttpContext context)
    {
        context.Response.ContentType = "text/plain";
        context.Response.Write("Hello World");
        Process(context);

    }
    public void Process(HttpContext context)
    {
        var MyDocxTitle = "testingdoc";
        // Create Stream
        using (MemoryStream mem = new MemoryStream())
        {
            // Create Document
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body docBody = new Body();

                // Add your docx content here
            }

            // Download File
            context.Response.AppendHeader("Content-Disposition", String.Format("attachment;filename=\"{0}.docx\"", MyDocxTitle));
            mem.Position = 0;
            mem.CopyTo(context.Response.OutputStream);
            context.Response.Flush();
            context.Response.End();
        }
    }
    public bool IsReusable
    {
        get
        {
            return false;
        }
    }

}