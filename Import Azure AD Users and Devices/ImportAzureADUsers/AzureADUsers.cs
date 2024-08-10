using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using Microsoft.IdentityModel.Protocols;
using Microsoft.Graph;

namespace ImportAzureADUsers
{
    internal class AzureADUsers
    {


        string clientID, tenantID, emailAddress, emailPassword;

        public string thisClassName()
        {
            return "This classs name is AzureADUSers";
        
        }

        public string getClientID()
        {

        
            try
            {
                return ConfigurationManager.ConnectionStrings["ClientID"].ToString();
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
             
        }


        public  async Task<string> getAzureADUsers()
         {

            try
            {


                clientID = ConfigurationManager.ConnectionStrings["ClientID"].ToString();
                tenantID = ConfigurationManager.ConnectionStrings["TenantID"].ToString();
                emailAddress= ConfigurationManager.ConnectionStrings["EmailAddress"].ToString();
                emailPassword   = ConfigurationManager.ConnectionStrings["EmailPassword"].ToString();




                MGraphServices graphServices = new MGraphServices();
                GraphServiceClient graphServiceClient =await  graphServices.connectToAzureAccount(clientID,
                           tenantID,
                           emailAddress,
                           emailPassword
                           );

                var users = await graphServiceClient.Users.Request().GetAsync();




                return "OK";
            }
            catch (Exception ex)
            {


                return ex.Message.ToString();
            }

        }








    }
}
