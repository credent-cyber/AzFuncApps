using System;
using System.IO;
using Demo.ITextSharp;

namespace Demo.ITextSharp.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            SignatureHelper.Sign("sample.pdf",
                "sample-signed.pdf",
                @"D:\Azure Certifications\az204\Tutorials\myfunc\Certificates\PnP.Core.SDK.AzureFunctionSample.cer",
                "no reason",
                "noida"
                );

            //byte[] contents;
            //int offset = 0;
            //using(var ims = new FileStream("sample.pdf", FileMode.Open))
            //{
            //    contents = new byte[ims.Length];
            //    while (true){
            //        int bytesRead = ims.Read(contents, 0, contents.Length);

            //        if (bytesRead <= 0)
            //            break;

            //        offset += bytesRead;
            //    }
            //}

            //var oms = new MemoryStream();
            //    SignatureHelper.Sign(contents,
            //           @"D:\Azure Certifications\az204\Tutorials\myfunc\Certificates\PnP.Core.SDK.AzureFunctionSample.cer",
            //           "no reason",
            //           "noida",
            //           ref oms
            //           );

            //oms.Seek(0, SeekOrigin.Begin);
            //using (var fs = new FileStream("sample-out-stream.pdf", FileMode.Create)) { 
            //    byte[] buffer = new byte[1024];

            //    while (oms.Read(buffer, 0, buffer.Length) > 0)
            //    {
            //        fs.Write(buffer, 0, buffer.Length);
            //    }
            //};

        }
    }
}
