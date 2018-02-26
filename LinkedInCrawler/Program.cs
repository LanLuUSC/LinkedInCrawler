namespace TestCrawler
{
    using System;
    using System.Net;
    using System.IO;
    using System.Text.RegularExpressions;
    using System.Diagnostics;
    using System.Threading;

    using HtmlAgilityPack;
    using OfficeOpenXml;


    class Program
    {
        public static bool AddCookies(HttpWebRequest request)// add fresh cookies here 
        {
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(new Cookie("JSESSIONID", "ajax:2564623502587446404", "/", "www.linkedin.com"));
            request.CookieContainer.Add(new Cookie("bscookie", "v=1&20160905235445fdeb9435-4510-486c-8daa-75551ecfe0d6AQF52nibqrT5Ef1gNbRj_7hrUe_pPEnm", "/", "www.linkedin.com"));
            request.CookieContainer.Add(new Cookie("oz_props_fetch_size1_undefined", "undefined", "/", "www.linkedin.com"));
            request.CookieContainer.Add(new Cookie("share_setting", "PUBLIC", "/", "www.linkedin.com"));
            request.CookieContainer.Add(new Cookie("sl", "v=1&y2i09", "/", "www.linkedin.com"));
            request.CookieContainer.Add(new Cookie("li_at", "AQEDAR6nbwsFLhlVAAABVvzGsJAAAAFW_TSNkEsAo6_TbuVwyp0PSx0zX3OCQwArS8YYp8XqCYa2iWo_QPojGdNH8QFULto5EOAJuARHsgVjREHnHMOGX3nkXqiG0BgcvXgfoenkXWWJkvUzAmnhLsW9", "/", "www.linkedin.com"));
            request.CookieContainer.Add(new Cookie("visit", "v=1&M", "/", "www.linkedin.com"));
            request.CookieContainer.Add(new Cookie("wutan", "4Bs3nkrPOlkA0TRdkIVzgOMg0naKYKgabla77TQjaXY=", "/", "www.linkedin.com"));

            return true;

        }

        /// <summary>
        /// Extracts from linkedin. First send a request to Google Search, with firstName, last name, email address and company name to get most possible linkedinURL(the first entry in google search)
        /// then add cookies(from a test linkedin account) and send linkedIn URL to linkedIn, get profile and extract information from it 
        /// </summary>
        /// <param name="firstName">The first name.</param>
        /// <param name="lastName">The last name.</param>
        /// <param name="emailAddress">The email address.</param>
        /// <param name="companyName">Name of the company.</param>
        /// <param name="title">The title.</param>
        /// <param name="url">The URL.</param>
        /// <param name="organizationName">Name of the organization.</param>
        /// <returns></returns>
        private static bool ExtractFromLinkedin(string firstName, string lastName, string emailAddress, string companyName, out string title, out string url, out string organizationName, out string location)
        {
            title = string.Empty;
            url = string.Empty;
            organizationName = string.Empty;
            location = string.Empty;

            string googleRequestInfo = "http://www.google.com/search?q=";
            if (!string.IsNullOrEmpty(firstName))
            {
                googleRequestInfo += firstName.ToLower();
            }

            if (!string.IsNullOrEmpty(lastName))
            {
                googleRequestInfo += "+" + lastName.ToLower();
            }

            
            if (!string.IsNullOrEmpty(emailAddress))
            {
                string emailDomain = Regex.Match(emailAddress, @"[\w\.\-]+@(.+)\.\w+").Groups[1].Value;
                googleRequestInfo += "+" + emailDomain;
            }

            if (!string.IsNullOrEmpty(companyName))
            {
                googleRequestInfo += "+" + companyName;
            }

            googleRequestInfo += "+" + "linkedin";
            HtmlWeb googleHtmlWeb = new HtmlWeb();
            googleHtmlWeb.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36";
            var googleReceivedDocument = googleHtmlWeb.Load(googleRequestInfo);

            url = string.Empty;

            var candidateNodes = googleReceivedDocument.DocumentNode.SelectNodes("//h3[@class='r']");
            if (candidateNodes != null)
            {
                var node = candidateNodes[0];
                string urlhref = string.Empty;
                string resultTitle = string.Empty;
                var child = node.FirstChild;
                if (child == null) return false;
                url = child.GetAttributeValue("href", string.Empty);
                if (url != string.Empty) resultTitle = child.InnerText;
                else return false;
                if (!url.Contains("https://www.linkedin.com/in/")) return false;

            }
            else return false;


            //connect to LinkedIn Part
            //HttpWebRequest linkedinRequest = (HttpWebRequest)WebRequest.Create("https://www.linkedin.com/in/lan-lu-a989788b");
            string linkedinRequestInfo = url;

            HtmlWeb linkedinHtmlWeb = new HtmlWeb();
            linkedinHtmlWeb.PreRequest = AddCookies;//Add cookies delegate
            linkedinHtmlWeb.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36";
            var linkedinReceivedDocument = linkedinHtmlWeb.Load(linkedinRequestInfo);
            candidateNodes = linkedinReceivedDocument.DocumentNode.SelectNodes("//div[@class='editable-item section-item current-position']");


            var locationNodes = linkedinReceivedDocument.DocumentNode.SelectNodes("//a[@name='location']");
            if (locationNodes != null)
            {
                location = locationNodes[0].InnerText;
            }

            if (candidateNodes != null)
            {
                foreach (var node in candidateNodes)
                {
                    var header = node.FirstChild?.FirstChild;
                    if (header == null) return false;
                    bool titleVisted = false;
                    foreach (var subNode in header.ChildNodes)
                    {
                        string s = subNode.InnerText.Trim();
                        if (string.IsNullOrEmpty(s)) continue;
                        if (!titleVisted)
                        {
                            titleVisted = true;
                            title = s;
                        }
                        else
                        {
                            organizationName = s;
                            break;
                        }
                    }

                }
            }
            else return false;


            return true;

        }


        static void Main(string[] args)
        {
            //change excel name here
            string testXlsx = @"C:\Users\Ryan\Desktop\Salesforce_Investors.xlsx";
            var file = new FileInfo(testXlsx);


            //string firstName = "jiuyang";
            //string lastName = "zhao";
            //string emailAddress = "";
            //string companyName = "columbia";
            string firstName;
            string lastName;
            string emailAddress;
            string companyName;
            string location;

            //string title, url, organizationName;

            //ExtractFromLinkedin(firstName, lastName, emailAddress, companyName, out title, out url, out organizationName);

            using (var excelPackage = new ExcelPackage(file))
            {
                float failCount = 0;
                float totalCount = 0;
                var randomGenerator = new Random();
                try
                {
                    var worksheet = excelPackage.Workbook.Worksheets[1];
                    int rowCount = worksheet.Dimension.End.Row;



                    for (int i = 1443; i <= 1590; ++i)
                    {
                        System.Console.WriteLine("Start to process {0} row", i - 1);

                        firstName = worksheet.Cells[i, 1].Value?.ToString();
                        lastName = worksheet.Cells[i, 2].Value?.ToString();
                        emailAddress = worksheet.Cells[i, 7].Value?.ToString();
                        companyName = worksheet.Cells[i, 3].Value?.ToString();

                        string title, url, organizationName;
                        totalCount = i - 1;
                        var watch = Stopwatch.StartNew();

                        if (ExtractFromLinkedin(firstName, lastName, emailAddress, companyName, out title, out url, out organizationName, out location))
                        {
                            worksheet.Cells[i, 4].Value = title;
                            if (string.IsNullOrEmpty(worksheet.Cells[i, 3].Value?.ToString()))
                            {
                                worksheet.Cells[i, 3].Value = organizationName;
                            }
                            worksheet.Cells[i, 20].Value = url;
                            if (!string.IsNullOrEmpty(location))
                            {
                                worksheet.Cells[i, 21].Value = location;
                            }

                            int sleeptime = randomGenerator.Next() % 300000;
                            System.Threading.Thread.Sleep(sleeptime);
                        }
                        else
                        {
                            worksheet.Cells[i, 4].Value = "Not enough information";
                            System.Console.WriteLine("Fail");
                            ++failCount;
                        }
                        watch.Stop();

                        var elapsedMs = watch.ElapsedMilliseconds;


                        System.Console.WriteLine("End to process {0} row, time cost {1}", i - 1, elapsedMs);

                    }


                }
                catch (Exception e)
                {
                    System.Console.WriteLine("Fatal Error in row{0}\t {1}", totalCount, e.Message);

                }
                finally
                {

                    float failRate = failCount / totalCount;
                    System.Console.WriteLine("Failure rate is {0}", failRate);
                    excelPackage.Save();


                }


            }

            System.Console.ReadKey();

        }
    }
}
