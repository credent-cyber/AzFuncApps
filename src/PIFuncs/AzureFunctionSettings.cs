﻿using System.Security.Cryptography.X509Certificates;

namespace Demo
{
    class AzureFunctionSettings
    {
        public string SiteUrl { get; set; }
        public string TestPortal { get; set; }
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public StoreName CertStoreName { get; set; }
        public StoreLocation CertStoreLocation { get; set; }
        public string CertificateThumbPrint { get; set; }
        public string CertPath { get; set; }
        public string Pwd { get; set; }

    }
}
