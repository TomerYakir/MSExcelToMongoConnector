using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Graph;


namespace MSExcelToMongoConnector
{
    class MSGraphHelper
    {
        #region Authentication
        // get token for application
        public static string GetTokenForApplication()
        {
            string TenantName = "tomeryakir.onMicrosoft.com";
            string AuthString = "https://login.microsoftonline.com/" + TenantName;
            string ResourceUrl = "https://graph.windows.net";
            AuthenticationContext authenticationContext = new AuthenticationContext(AuthString, false);
            // Config for OAuth client credentials 
            // client ID - 35447a3d-03c1-4e02-a405-c60f2b142d41
        
            string clientId = "35447a3d-03c1-4e02-a405-c60f2b142d41";
            ClientCredential clientCredNoSecret = new ClientCredential(clientId, "");
            /*
             Task<AuthenticationResult> authenticationResult = await authenticationContext.AcquireTokenAsync(ResourceUrl,
                clientCredNoSecret);
                */
            string token = ""; // authenticationResult.;
            return token;
        }

        #endregion
    }
}
