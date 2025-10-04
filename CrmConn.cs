using System;
using System.Configuration;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;         // <— AJOUT
using Microsoft.Xrm.Tooling.Connector;

public static class CrmConn
{
    public static IOrganizationService Connect(out string diag)
    {
        diag = null;
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        var cs = ConfigurationManager.AppSettings["CrmConnectionString"]?.Trim();
        if (string.IsNullOrWhiteSpace(cs))
        {
            diag = "CrmConnectionString is missing in App.config";
            return null;
        }

        try
        {
            var client = new CrmServiceClient(cs);

            if (!client.IsReady)
            {
                diag = "LastCrmError: " + (client.LastCrmError ?? "(null)") + Environment.NewLine
                     + "Org: " + (client.ConnectedOrgUniqueName ?? "(none)") + Environment.NewLine
                     + "Svc: " + (client.CrmConnectOrgUriActual?.ToString() ?? "(none)"); // <— propriété correcte
                return null;
            }

            var svc = client.OrganizationWebProxyClient != null
                ? (IOrganizationService)client.OrganizationWebProxyClient
                : client.OrganizationServiceProxy;

            // Vérification immédiate sans dépendre des types WhoAmI*
            var whoReq = new Microsoft.Xrm.Sdk.OrganizationRequest("WhoAmI");
            var whoResp = svc.Execute(whoReq); // OrganizationResponse
            var userId = (System.Guid)whoResp.Results["UserId"];
            if (userId == System.Guid.Empty)
            {
                diag = "Connected but WhoAmI returned empty UserId";
                return null;
            }

            return svc;
        }
        catch (Exception ex)
        {
            diag = "EX: " + ex.Message + Environment.NewLine
                 + "INNER: " + ex.InnerException?.Message;
            return null;
        }
    }
}