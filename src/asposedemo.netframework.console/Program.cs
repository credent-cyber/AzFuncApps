using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace asposedemo.netframework.console
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var opt = new Aspose.CAD.LoadOptions();
            //var file = @"D:\Downloads\sample1.dwg";
            var file = @"D:\Downloads\1344464555 (1).dwg"; // 12mb
            //var file = @"D:\Downloads\sample1 (2).dwg";
            var outfile = "out.pdf";
            var watch = Stopwatch.StartNew();
            var running = true;
            var task = Task.Factory.StartNew(() => { 
                while (running) {
                    Thread.Sleep(1000);
                    Console.Clear();
                    Console.WriteLine($"Elapsed seconds : {(int)watch.Elapsed.TotalSeconds}"); 
                } 
            });
            
            using (var img = Aspose.CAD.Image.Load(file, opt))
            {
                var w = img.Width;
                var h = img.Height;

                Aspose.CAD.ImageOptions.CadRasterizationOptions rasterizationOptions = new Aspose.CAD.ImageOptions.CadRasterizationOptions();
                rasterizationOptions.PageWidth = w;
                rasterizationOptions.PageHeight = h;
                rasterizationOptions.AutomaticLayoutsScaling = true;
                rasterizationOptions.NoScaling = false;

                // Create an instance of PdfOptions
                Aspose.CAD.ImageOptions.PdfOptions pdfOptions = new Aspose.CAD.ImageOptions.PdfOptions();
                
                //Set the VectorRasterizationOptions property
                pdfOptions.VectorRasterizationOptions = rasterizationOptions;
                
                using(Stream os = new FileStream(outfile, FileMode.Create, FileAccess.Write, FileShare.None))
                //Export CAD to PDF
                img.Save(os, pdfOptions);

            }
            watch.Stop();
            running = false;
            Console.WriteLine($"Total Time taken (in sec) - {watch.Elapsed.TotalSeconds}, press any key contine");
            Console.ReadKey();
        }
    }
}
