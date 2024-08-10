using Microsoft.Graph;
using System;
using System.Configuration;
using Microsoft.IdentityModel.Protocols;
using System.Threading.Tasks;

namespace ImportAzureADUsers
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            //THE BELOW ARE USED FOR AZURE APP AND EMAIL CREDENTIALS
            string clientID, tenantID, emailAddress, emailPassword, errMessage;
            
            //BELOW ARE USED TO STORE DEVICE DETAILS
            string deviceID,     disaplyName,  lastSignInDate,    operatingSystem,    operatingSystemVersion,   deviceMetaDatada;
            
            //BELOW ARE USED TO STORE AD USER DETAILS
            string userPrincipalName, userEmail, userSurname ,  userCompanyName , userStreetAddress, userCity, userState, userCountry , userJobTitle;


           


            try
            {

                //BELOW LINES READ TEH CREDENTIALS FROM APP.CONFIG FILE. (PLEASE ADD APPROPRIATE CREDENTIALS THERE)
                clientID = ConfigurationManager.ConnectionStrings["ClientID"].ToString();
                tenantID = ConfigurationManager.ConnectionStrings["TenantID"].ToString();
                emailAddress = ConfigurationManager.ConnectionStrings["EmailAddress"].ToString();
                emailPassword = ConfigurationManager.ConnectionStrings["EmailPassword"].ToString();




                MGraphServices graphServices = new MGraphServices();
                GraphServiceClient graphServiceClient = await graphServices.connectToAzureAccount(clientID,
                tenantID,
                           emailAddress,
                           emailPassword
                           );


                //BELOW LINE IS USED TO READ USERS FROM AZURE AD
                var ADusers = await graphServiceClient.Users.Request().Expand(m => m.Manager)
               .Select("id,companyName,streetAddress,city,postalCode,state,country,department,businessPhones,displayName,jobTitle,mail,mobilePhones,userType,employeeHireDate,employeeHireDate,externalEmployeeID,ageGroup,consentProvidedForMinor,legalageGroupClassification,employeeId,accountEnabled,employeeType")
                       .GetAsync();

                foreach (var ADuser in ADusers)
                {

                    userPrincipalName = ADuser.UserPrincipalName;
                    userEmail = ADuser.Mail;
                    userSurname= ADuser.Surname;
                    userCompanyName= ADuser.CompanyName;
                    userStreetAddress= ADuser.StreetAddress;
                    userCity= ADuser.City;
                    userState= ADuser.State;
                    userCountry=  ADuser.Country;
                    userJobTitle= ADuser.JobTitle;

                }

                //BELOW LINE IS USED TO READ THE DEVICES
                var devices = await graphServiceClient.Devices.Request().Expand(d => d.RegisteredOwners).GetAsync();


                foreach (var device in devices)
                {
                   
                    deviceID = device.DeviceId;
                    disaplyName = device.DisplayName;
                    lastSignInDate = device.ApproximateLastSignInDateTime.ToString();
                    operatingSystem = device.OperatingSystem;
                    operatingSystemVersion = device.OperatingSystemVersion;
                    deviceMetaDatada = device.DeviceMetadata;

                }




                }
            catch (Exception ex)
            {
                        errMessage= ex.Message.ToString();
            }


        }
    }
}
