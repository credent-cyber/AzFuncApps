using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Signatures;
using Org.BouncyCastle.X509;
using System;
using System.IO;
using System.Net;

namespace Demo.ITextSharp
{
    public class SignatureHelper
    {
        /// <summary>
        /// To sign a local document
        /// </summary>
        /// <param name="src"></param>
        /// <param name="dest"></param>
        /// <param name="certFile"></param>
        /// <param name="reason"></param>
        /// <param name="location"></param>
        public static void Sign(string src, string dest, string certFile, string reason, string location)
        {
            //var src = "sample.pdf";
            var pdf = new PdfReader(src);
            var signer = new PdfSigner(pdf,
                new FileStream(dest, FileMode.Create),
                new StampingProperties()
                );

            Rectangle rect = new Rectangle(36, 648, 200, 100);
            PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
            appearance
                .SetReason(reason)
                .SetLocation(location)
                .SetPageRect(rect)
                .SetPageNumber(1);
            signer.SetFieldName("sig");

            var externalSignature = new AuthoritySignature();

            //var certFile = @"D:\Azure Certifications\az204\Tutorials\myfunc\Certificates\PnP.Core.SDK.AzureFunctionSample.cer";
            using (var stream = new FileStream(certFile, FileMode.Open))
            {
                X509CertificateParser parser = new X509CertificateParser();
                X509Certificate[] chain = new X509Certificate[1];
                chain[0] = parser.ReadCertificate(stream);
                signer.SignDetached(externalSignature, chain, null, null, null, 0, PdfSigner.CryptoStandard.CMS);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="content"></param>
        /// <param name="dest"></param>
        /// <param name="certFile"></param>
        /// <param name="reason"></param>
        /// <param name="location"></param>
        /// <param name="outStream"></param>
        public static void Sign(byte[] content, string dest, string certFile, string reason, string location)
        {
            using (var stream = new MemoryStream(content))
            {
                var pdf = new PdfReader(stream);

                var signer = new PdfSigner(pdf,
                                   new FileStream(dest, FileMode.Create),
                                   new StampingProperties()
                           );

                Rectangle rect = new Rectangle(36, 648, 200, 100);
                PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
                appearance
                    .SetReason(reason)
                    .SetLocation(location)
                    .SetPageRect(rect)
                    .SetPageNumber(1);
                signer.SetFieldName("sig");

                var externalSignature = new AuthoritySignature();

                //var certFile = @"D:\Azure Certifications\az204\Tutorials\myfunc\Certificates\PnP.Core.SDK.AzureFunctionSample.cer";
                using (var certStream = new FileStream(certFile, FileMode.Open))
                {
                    X509CertificateParser parser = new X509CertificateParser();
                    X509Certificate[] chain = new X509Certificate[1];
                    chain[0] = parser.ReadCertificate(certStream);
                    signer.SignDetached(externalSignature, chain, null, null, null, 0, PdfSigner.CryptoStandard.CMS);
                    var doc = signer.GetDocument();
                }
            }
        }



    }

    public class AuthoritySignature : IExternalSignature
    {
        public static readonly string SIGN = "CREDENT";

        public string GetEncryptionAlgorithm()
        {
            return "RSA";
        }

        public string GetHashAlgorithm()
        {
            return DigestAlgorithms.SHA256;
        }

        public byte[] Sign(byte[] message)
        {
            return System.Text.Encoding.UTF8.GetBytes(SIGN);
        }
    }
}
