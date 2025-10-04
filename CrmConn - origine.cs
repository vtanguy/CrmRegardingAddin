using System.Configuration;
using System;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;

namespace CrmRegardingAddin
{
    public static class CrmConn
    {
        public static IOrganizationService Connect()
        {
            var connStr = System.Configuration.ConfigurationManager.AppSettings["CrmConnectionString"];
            if (string.IsNullOrWhiteSpace(connStr))
                throw new Exception("CrmConnectionString manquante dans App.config");
            var svc = new CrmServiceClient(connStr);
            if (!svc.IsReady)
                throw new Exception(svc.LastCrmError ?? "Connexion CRM échouée");
            return svc.OrganizationWebProxyClient ?? (IOrganizationService)svc.OrganizationServiceProxy;
        }
    }
}
