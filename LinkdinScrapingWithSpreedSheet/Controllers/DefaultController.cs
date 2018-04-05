using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using Spring.Json;
using Spring.Social.OAuth1;
using Spring.Social.LinkedIn.Api;
using Spring.Social.LinkedIn.Connect;
using System.Diagnostics;

using System.Collections.Specialized;
using System.Configuration;

namespace LinkdinScrapingWithSpreedSheet.Controllers
{
    public class DefaultController : Controller
    {
        // GET: Default
        public static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public static string ApplicationName = "Google Sheets API .NET Quickstart";

        public static string LinkedInApiKey = ConfigurationManager.AppSettings["LinkedInApiKey"];
        public static string LinkedInApiSecret = ConfigurationManager.AppSettings["LinkedInApiSecret"];
        public static string SheetID = ConfigurationManager.AppSettings["SheetID"];
        OAuthToken oauthToken = null;
        LinkedInServiceProvider linkedInServiceProvider = new LinkedInServiceProvider(LinkedInApiKey, LinkedInApiSecret);
        public ActionResult Index()
        {


            return View();
        }

        public ActionResult StartScraping()
        {
            NameValueCollection parameters = new NameValueCollection();

            parameters.Add("scope", "r_basicprofile r_emailaddress");
            var siteUrl = ConfigurationManager.AppSettings["siteurl"].ToString() + "Default/AutenticationURL";
            oauthToken = linkedInServiceProvider.OAuthOperations.FetchRequestTokenAsync(siteUrl, parameters).Result;
            Console.WriteLine("Done");

            string authenticateUrl = linkedInServiceProvider.OAuthOperations.BuildAuthorizeUrl(oauthToken.Value, null);
            Console.WriteLine("Redirect user for authentication: " + authenticateUrl);
            Session["RequestToken"] = oauthToken;


            return Redirect(authenticateUrl);


        }


        List<LinkdinCloneData> GetData()
        {
            UserCredential credential;
            string path = Server.MapPath("~/").ToString() + "client_secret.json";
            string Newpath = Server.MapPath("~/").ToString();
            using (var stream =
                new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                string credPath = Newpath + "/ sheets.googleapis.com-dotnet-quickstart.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = SheetID;
            String range = "Apps";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);


            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;


            if (values != null && values.Count > 1)
            {
                List<LinkdinCloneData> lcd = new List<LinkdinCloneData>();
                var i = 0;
                foreach (var row in values)
                {
                    if (i == 0)
                    {


                    }
                    else
                    {
                        LinkdinCloneData objlcd = new LinkdinCloneData();
                        if (row.Count > 0)
                            objlcd.FirstName = row[0].ToString();
                        if (row.Count > 1)
                            objlcd.FullName = row[1].ToString();
                        if (row.Count > 2)
                            objlcd.EmailAddress = row[2].ToString();
                        if (row.Count > 3)
                            objlcd.Wildcard = row[3].ToString();
                        if (row.Count > 4)
                            objlcd.Headline = row[4].ToString();
                        if (row.Count > 5)
                            objlcd.School = row[5].ToString();
                        if (row.Count > 6)
                            objlcd.Job = row[6].ToString();
                        if (row.Count > 7)
                            objlcd.summary = row[7].ToString();
                        if (row.Count > 8)
                            objlcd.AllPostions = row[8].ToString();
                        if (row.Count > 9)
                            objlcd.GraduationYear = row[9].ToString();
                        if (row.Count > 10)
                            objlcd.LinkedInURL = row[10].ToString();
                        if (row.Count > 11)
                            objlcd.App = row[11].ToString();
                        if (row.Count > 12)
                        {
                            objlcd.FileAttachments = row[12].ToString();
                        }
                        if (row.Count > 13)
                        {
                            objlcd.ScheduledDate = row[13].ToString();
                        }
                        if (row.Count > 14)
                        {
                            objlcd.MailMergeStatus = row[14].ToString();
                        }
                        if (row.Count > 15)
                        {
                            objlcd.CODE = row[15].ToString();
                        }
                        lcd.Add(objlcd);

                    }

                    i++;
                }
                return lcd;

            }
            else
            {


                return null;
            }
        }

        public ActionResult AutenticationURL(string oauth_verifier)
        {

            OAuthToken oauth_token = Session["RequestToken"] as OAuthToken;
            AuthorizedRequestToken requestToken = new AuthorizedRequestToken(oauth_token, oauth_verifier);
            OAuthToken oauthAccessToken = linkedInServiceProvider.OAuthOperations.ExchangeForAccessTokenAsync(requestToken, null).Result;

            ILinkedIn linkedIn = linkedInServiceProvider.GetApi(oauthAccessToken.Value, oauthAccessToken.Secret);
            var lst = GetData();

            foreach (var item in lst)
            {
                if (!string.IsNullOrEmpty(item.LinkedInURL))
                {
                    LinkedInFullProfile profileByPublicUrl = linkedIn.ProfileOperations.GetUserFullProfileByPublicUrlAsync(item.LinkedInURL).Result;
                    item.FullName = profileByPublicUrl.FirstName + " " + profileByPublicUrl.LastName;
                    item.FirstName = profileByPublicUrl.FirstName;
                    item.Headline = profileByPublicUrl.Headline;
                    var AllPostion = "";
                    for (int i = 0; i < profileByPublicUrl.Positions.Count; i++)
                    {
                        if (i == 0)
                        {
                            item.Job = profileByPublicUrl.Positions[i].Title + "-" + profileByPublicUrl.Positions[i].Company.Name;
                            AllPostion = profileByPublicUrl.Positions[i].Title + "-" + profileByPublicUrl.Positions[i].Company.Name;
                        }
                        else
                        {
                            AllPostion += ",  " + profileByPublicUrl.Positions[i].Title + "-" + profileByPublicUrl.Positions[i].Company.Name;
                        }
                    }
                    item.summary = profileByPublicUrl.Summary;
                    item.AllPostions = AllPostion;


                }
            }

            ///update Google Sheet





            ViewBag.LinkdinScrpedList = lst;
            var service = AuthorizeGoogleApp();
            UpdateGoogleRow(service, lst);
            return View();
        }
        private  SheetsService AuthorizeGoogleApp()
        {
            UserCredential credential;
            string Newpath = Server.MapPath("~/").ToString();
            string path = Server.MapPath("~/").ToString() + "client_secret.json";
            using (var stream =
                new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                string credPath = Newpath + "/ sheets.googleapis.com-dotnet-quickstart.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }
        private static void UpdateGoogleRow(SheetsService service, List<LinkdinCloneData> lst)
        {
            List<IList<Object>> objNewRecords = new List<IList<Object>>();
            foreach (var item in lst)
            {
                IList<Object> obj = new List<Object>();

                obj.Add(item.FirstName);
                obj.Add(item.FullName);
                obj.Add(item.EmailAddress);
                obj.Add(item.Wildcard);
                obj.Add(item.Headline);
                obj.Add(item.School);
                obj.Add(item.Job);
                obj.Add(item.summary);
                obj.Add(item.AllPostions);

                obj.Add(item.GraduationYear);
                obj.Add(item.LinkedInURL);
                obj.Add(item.App);
                obj.Add(item.FileAttachments);
                obj.Add(item.ScheduledDate);
                obj.Add(item.MailMergeStatus);
                obj.Add(item.CODE);
                objNewRecords.Add(obj);
            }

            String spreadsheetId2 = SheetID;
            String range2 = "A2:P";  //         
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(new ValueRange() { Values = objNewRecords }, spreadsheetId2, range2);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();

            Console.WriteLine("done!");
        }
    }
    public class LinkdinCloneData
    {
        public string FirstName { get; set; }
        public string FullName { get; set; }
        public string EmailAddress { get; set; }
        public string Wildcard { get; set; }
        public string Headline { get; set; }
        public string School { get; set; }
        public string Job { get; set; }
        public string GraduationYear { get; set; }
        public string LinkedInURL { get; set; }
        public string App { get; set; }
        public string FileAttachments { get; set; }
        public string ScheduledDate { get; set; }
        public string MailMergeStatus { get; set; }
        public string CODE { get; set; }
        public string summary { get; set; }
        public string AllPostions { get; set; }
    }
}