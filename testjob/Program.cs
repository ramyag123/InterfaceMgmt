using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Microsoft.SharePoint.Client;
using System.IO;
using File = System.IO.File;
using System.Linq;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using Microsoft.ProjectServer.Client;
using System.Runtime.Remoting.Contexts;
using System.Threading;
namespace testjob
{
    public class TaskDetails
    {
        public string GUID { get; set; }
        public string Title { get; set; }
        public string TStartDate { get; set; }
        public string TActStartDate { get; set; }
        public string TFinishDate { get; set; }
        public string FDatePlusonee { get; set; }

        public override string ToString()
        {
            return $"[GUID: {GUID}, Title: {Title}, TStartDate: {TStartDate}, TFinishDate: {TFinishDate}, FDatePlusonee: {FDatePlusonee}]";
        }


    }
    public class PrjDetails
    {
        public string GUID { get; set; }
        public string wfInP { get; set; }
        public string PubByWf { get; set; }


        public override string ToString()
        {
            return $"[GUID: {GUID}, WFinProg: {wfInP}, Pubbywf: {PubByWf}]";
        }


    }


    class Program
    {
        private static readonly string logDirectory = "C:\\Users\\Pecs\\Downloads\\logs";
        //  private static readonly string logDirectory = "C:\\Users\\Ramya\\Downloads\\logs"; // Folder to store logs

        private static readonly string logFilePath;
        private static readonly HttpClient httpClient = new HttpClient(new HttpClientHandler());

        static Program()
        {
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            logFilePath = Path.Combine(logDirectory, $"ConsoleLog_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.txt");
        }

        static void LogMessage(string message)
        {
            try
            {
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}";
                File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
                Console.WriteLine(logEntry);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to log file: {ex.Message}");
            }
        }

        static async Task<string> GetRequestDigest(HttpClient httpClient, string siteUrl)
        {
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, $"{siteUrl}/_api/contextinfo");
            request.Headers.Add("Accept", "application/json;odata=verbose");

            HttpResponseMessage response = await httpClient.SendAsync(request);
            if (!response.IsSuccessStatusCode)
            {
                LogMessage($"❌ Failed to get X-RequestDigest: {response.StatusCode}");
                return null;
            }

            string content = await response.Content.ReadAsStringAsync();
            dynamic jsonResponse = JObject.Parse(content);
            return jsonResponse.d.GetContextWebInformation.FormDigestValue;
        }

        static async Task<(List<string> TaskGuids, List<string> ProjGuids)> FetchAllInterfaceManagementItems(HttpClient httpClient, string siteUrl, string listName)
        {
            HashSet<string> taskGuids = new HashSet<string>(); // ✅ Ensures uniqueness
            HashSet<string> projGuids = new HashSet<string>(); // ✅ Ensures uniqueness

            string digestValue = await GetRequestDigest(httpClient, siteUrl); // ✅ Refresh Digest
            Console.WriteLine($"🔄 Fresh X-RequestDigest: {digestValue}");

            string nextPageUrl = $"{siteUrl}/_api/web/lists/getbytitle('{listName}')/items?$top=500";
            int pageCount = 0;
            var i = 0;
            do
            {
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, nextPageUrl);
                request.Headers.Add("Accept", "application/json;odata=verbose");

                HttpResponseMessage response = await httpClient.SendAsync(request);

                if (!response.IsSuccessStatusCode)
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"❌ Failed to fetch items. Status: {response.StatusCode}, Error: {errorContent}");
                    break;
                }

                string responseContent = await response.Content.ReadAsStringAsync();
                JObject jsonResponse = JObject.Parse(responseContent);

                if (jsonResponse["d"]["results"] != null)
                {
                    foreach (var item in jsonResponse["d"]["results"])
                    {
                        i++;
                        string sourceGuid = item["Task_x0020_Name"]?.ToString();
                        string dependentGuid = item["Dependent_x0020_Task_x0020_Name"]?.ToString();

                        string prjguid = item["Project_x0020_Depends_x0020_on"]?.ToString();
                        if (!string.IsNullOrEmpty(sourceGuid)) taskGuids.Add(sourceGuid);
                        if (!string.IsNullOrEmpty(dependentGuid)) taskGuids.Add(dependentGuid);
                        if (!string.IsNullOrEmpty(prjguid)) projGuids.Add(prjguid);

                    }
                }

                // 🔹 Get the pagination URL
                nextPageUrl = jsonResponse["d"]["__next"]?.ToString();

                pageCount++;
                Console.WriteLine($"📌 Page {pageCount} processed. Next Page: {nextPageUrl ?? "No More Pages"}");

            } while (!string.IsNullOrEmpty(nextPageUrl)); // ✅ Continue until all pages are fetched

            Console.WriteLine($"✅ Total Task Unique GUIDs Retrieved: {taskGuids.Count} " + i);
            Console.WriteLine($"✅ Total Prj Unique GUIDs Retrieved: {projGuids.Count} " + i);

            return (taskGuids.ToList(), projGuids.ToList());
        }

        /*public static async Task<List<TaskDetails>> FetchProjectTasks(HttpClient httpClient, string siteUrl, List<string> taskGuids)
        {
            List<TaskDetails> taskList = new List<TaskDetails>();
            int batchSize = 2; // OData allows around 200 conditions per query

            //  for (int i = 0; i < taskGuids.Count; i += batchSize)
            // {
            // var batchGuids = taskGuids.Skip(i).Take(batchSize);
            var batchGuids = taskGuids.Take(30);

            string filterCondition = string.Join(" or ", batchGuids.Select(g => $"TaskId eq guid'{g}'"));

                string requestUrl = $"{siteUrl}/_api/ProjectData/Tasks?$filter={filterCondition}&$select=TaskId,TaskStartDate,TaskActualStartDate,TaskWBS,TaskName,TaskFinishDate,FinishDatePlusonee";
                Console.WriteLine(requestUrl);
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                request.Headers.Add("Accept", "application/json;odata=verbose");

                HttpResponseMessage response = await httpClient.SendAsync(request);
                string responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"❌ Error fetching tasks: {response.StatusCode}, Response: {responseContent}");
                // continue;
                }

                JObject jsonResponse = JObject.Parse(responseContent);
                var items = jsonResponse["d"]["results"];

                foreach (var item in items)
                {

                taskList.Add(new TaskDetails
                {
                    GUID = item["TaskId"].ToString(),
                    Title = item["TaskName"].ToString(),
                    TStartDate = item["TaskStartDate"].ToString(),
                    TFinishDate = item["TaskFinishDate"].ToString(),
                    FDatePlusonee = item["FinishDatePlusonee"].ToString()

                });

            }

            // Console.WriteLine($"✅ Retrieved {items.Count()} tasks from batch {i / batchSize + 1}");
            //}

            Console.WriteLine($"🎯 Total Tasks Retrieved: {taskList.Count}");
            return taskList;
        }*/
        public static async Task<List<TaskDetails>> FetchProjectTasks(HttpClient httpClient, string siteUrl, List<string> taskGuids)
        {
            List<TaskDetails> taskList = new List<TaskDetails>();
            int batchSize = 30; // Process 30 tasks at a time
            HashSet<string> fetchedGuids = new HashSet<string>(); // To track retrieved GUIDs
            taskGuids = taskGuids.Where(g => g != "0" && g != "00000000-0000-0000-0000-000000000000").ToList();

            for (int i = 0; i < taskGuids.Count; i += batchSize)
            {
                var batchGuids = taskGuids.Skip(i).Take(batchSize);
                string filterCondition = string.Join(" or ", batchGuids.Select(g => $"TaskId eq guid'{g}'"));

                string requestUrl = $"{siteUrl}/_api/ProjectData/Tasks?$filter={filterCondition}&$select=TaskId,TaskStartDate,TaskActualStartDate,TaskWBS,TaskName,TaskFinishDate,FinishDatePlusonee";
                Console.WriteLine($"🔍 Fetching batch {i / batchSize + 1}: {requestUrl}");

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                request.Headers.Add("Accept", "application/json;odata=verbose");

                HttpResponseMessage response = await httpClient.SendAsync(request);
                string responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"❌ Error fetching batch {i / batchSize + 1}: {response.StatusCode}, Response: {responseContent}");
                    continue; // Skip this batch
                }

                JObject jsonResponse = JObject.Parse(responseContent);
                var items = jsonResponse["d"]["results"];

                foreach (var item in items)
                {
                    string taskId = item["TaskId"].ToString();
                    fetchedGuids.Add(taskId);

                    taskList.Add(new TaskDetails
                    {
                        GUID = taskId,
                        Title = item["TaskName"].ToString(),
                        TStartDate = item["TaskStartDate"]?.ToString(),
                        TFinishDate = item["TaskFinishDate"]?.ToString(),
                        FDatePlusonee = item["FinishDatePlusonee"]?.ToString()
                    });
                }

                Console.WriteLine($"✅ Retrieved {items.Count()} tasks in batch {i / batchSize + 1}");
                await System.Threading.Tasks.Task.Delay(500);
            }

            // 🔄 Check for Missing Tasks
            var missingGuids = taskGuids.Except(fetchedGuids).ToList();
            if (missingGuids.Count > 0)
            {
                Console.WriteLine($"⚠️ Missing {missingGuids.Count} tasks, fetching individually...");

                foreach (var guid in missingGuids)
                {
                    string singleTaskUrl = $"{siteUrl}/_api/ProjectData/Tasks?$filter=TaskId eq guid'{guid}'";
                    Console.WriteLine($"🔍 Fetching Task: {singleTaskUrl}");

                    HttpResponseMessage singleResponse = await httpClient.GetAsync(singleTaskUrl);
                    string singleContent = await singleResponse.Content.ReadAsStringAsync();

                    if (!singleResponse.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"❌ Failed to fetch Task {guid}: {singleResponse.StatusCode}, Response: {singleContent}");
                        continue;
                    }

                    // 🛑 Handle Unexpected HTML Responses
                    if (singleContent.Trim().StartsWith("<"))
                    {
                        Console.WriteLine($"⚠️ Received HTML response for Task {guid}, skipping...");
                        continue;
                    }

                    JObject singleJsonResponse = JObject.Parse(singleContent);
                    var singleItem = singleJsonResponse["d"]["results"].FirstOrDefault();
                    if (singleItem != null)
                    {
                        taskList.Add(new TaskDetails
                        {
                            GUID = singleItem["TaskId"].ToString(),
                            Title = singleItem["TaskName"].ToString(),
                            TStartDate = singleItem["TaskStartDate"]?.ToString(),
                            TActStartDate = singleItem["TaskActualStartDate"]?.ToString(),
                            TFinishDate = singleItem["TaskFinishDate"]?.ToString(),
                            FDatePlusonee = singleItem["FinishDatePlusonee"]?.ToString()
                        });
                    }

                    await System.Threading.Tasks.Task.Delay(500);
                }
            }

            Console.WriteLine($"🎯 Total Tasks Retrieved: {taskList.Count}");
            return taskList;
        }

        public static async Task<List<PrjDetails>> FetchDProject(HttpClient httpClient, string siteUrl, List<string> prjGuids)
        {
            List<PrjDetails> prjList = new List<PrjDetails>();
            int batchSize = 30; // Process 30 tasks at a time
            HashSet<string> fetchedGuids = new HashSet<string>(); // To track retrieved GUIDs
            prjGuids = prjGuids.Where(g => g != "0" && g != "00000000-0000-0000-0000-000000000000").ToList();

            for (int i = 0; i < prjGuids.Count; i += batchSize)
            {
                var batchGuids = prjGuids.Skip(i).Take(batchSize);
                string filterCondition = string.Join(" or ", prjGuids.Select(g => $"ProjectId eq guid'{g}'"));
                string requestUrl = $"{siteUrl}/_api/ProjectData/Projects?$filter={filterCondition}";

                // string requestUrl = $"{siteUrl}/_api/ProjectData/Projects?$filter={filterCondition}&$select=Id,WorkflowInProgress,PublishedByWorkflow,TaskWBS,TaskName,TaskFinishDate,FinishDatePlusonee";
                Console.WriteLine($"🔍 Fetching projects batch {i / batchSize + 1}: {requestUrl}");

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                request.Headers.Add("Accept", "application/json;odata=verbose");

                HttpResponseMessage response = await httpClient.SendAsync(request);
                string responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"❌ Error fetching batch {i / batchSize + 1}: {response.StatusCode}, Response: {responseContent}");
                    continue; // Skip this batch
                }

                JObject jsonResponse = JObject.Parse(responseContent);
                var items = jsonResponse["d"]["results"];

                foreach (var item in items)
                {
                    string prjId = item["ProjectId"].ToString();
                    fetchedGuids.Add(prjId);
                    // return $"[GUID: {GUID}, WFinProg: {wfInP}, Pubbywf: {PubByWf}]";

                    prjList.Add(new PrjDetails
                    {
                        GUID = item["ProjectId"].ToString(),
                        wfInP = item["WorkflowInProgress"].ToString(),
                        PubByWf = item["PublishedByWorkflow"]?.ToString()
                    });
                }

                Console.WriteLine($"✅ Retrieved {items.Count()} tasks in batch {i / batchSize + 1}");
                await System.Threading.Tasks.Task.Delay(500);
            }

            // 🔄 Check for Missing Tasks
            var missingGuids = prjGuids.Except(fetchedGuids).ToList();
            if (missingGuids.Count > 0)
            {
                Console.WriteLine($"⚠️ Missing {missingGuids.Count} tasks, fetching individually...");

                foreach (var guid in missingGuids)
                {
                    string singlePrjUrl = $"{siteUrl}/_api/ProjectData/Projects/guid('{guid}')";
                    Console.WriteLine($"🔍 Fetching Project: {singlePrjUrl}");

                    HttpResponseMessage singleResponse = await httpClient.GetAsync(singlePrjUrl);
                    string singleContent = await singleResponse.Content.ReadAsStringAsync();

                    if (!singleResponse.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"❌ Failed to fetch Task {guid}: {singleResponse.StatusCode}, Response: {singleContent}");
                        continue;
                    }

                    // 🛑 Handle Unexpected HTML Responses
                    if (singleContent.Trim().StartsWith("<"))
                    {
                        Console.WriteLine($"⚠️ Received HTML response for Task {guid}, skipping...");
                        continue;
                    }

                    JObject singleJsonResponse = JObject.Parse(singleContent);
                    var singleItem = singleJsonResponse["d"]["results"].FirstOrDefault();
                    if (singleItem != null)
                    {
                        prjList.Add(new PrjDetails
                        {
                            GUID = singleItem["ProjectId"].ToString(),
                            wfInP = singleItem["WorkflowInProgress"].ToString(),
                            PubByWf = singleItem["PublishedByWorkflow"]?.ToString()
                        });

                    }

                    await System.Threading.Tasks.Task.Delay(500);
                }
            }

            Console.WriteLine($"🎯 Total Tasks Retrieved: {prjList.Count}");
            return prjList;
        }

        public static async Task<List<TaskDetails>> FetchProjectTasks892(HttpClient httpClient, string siteUrl, List<string> taskGuids)
        {
            List<TaskDetails> taskList = new List<TaskDetails>();
            int batchSize = 30; // ✅ Fetch 30 task GUIDs per request

            for (int i = 0; i < taskGuids.Count; i += batchSize)
            {
                var batchGuids = taskGuids.Skip(i).Take(batchSize);
                string filterCondition = string.Join(" or ", batchGuids.Select(g => $"TaskId eq guid'{g}'"));

                string requestUrl = $"{siteUrl}/_api/ProjectData/Tasks?$filter={filterCondition}&$select=TaskId,TaskStartDate,TaskActualStartDate,TaskWBS,TaskName,TaskFinishDate,FinishDatePlusonee";
                //Console.WriteLine($"🔍 Fetching batch {i / batchSize + 1}: {requestUrl}");

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                request.Headers.Add("Accept", "application/json;odata=verbose");

                HttpResponseMessage response = await httpClient.SendAsync(request);
                string responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"❌ Error fetching tasks: {response.StatusCode}, Response: {responseContent}");
                    continue; // ✅ Skip this batch and move to the next
                }

                JObject jsonResponse = JObject.Parse(responseContent);
                //  Console.WriteLine("🔍 API Response: " + responseContent); // ✅ Log full response

                JToken items = jsonResponse.SelectToken("d.results") ?? jsonResponse.SelectToken("value");

                if (items == null)
                {
                    Console.WriteLine("⚠️ No tasks found in response!");
                    continue; // ✅ Skip empty responses
                }

                foreach (var item in items)
                {
                    taskList.Add(new TaskDetails
                    {
                        GUID = item["TaskId"]?.ToString(),
                        Title = item["TaskName"]?.ToString(),
                        TStartDate = item["TaskStartDate"]?.ToString(),
                        TActStartDate = item["TaskActualStartDate"]?.ToString(),
                        TFinishDate = item["TaskFinishDate"]?.ToString(),
                        FDatePlusonee = item["FinishDatePlusonee"]?.ToString()
                    });
                }

                Console.WriteLine($"✅ Retrieved {items.Count()} tasks from batch {i / batchSize + 1}");

            }

            Console.WriteLine($"🎯 Total Tasks Retrieved: {taskList.Count}");
            return taskList;
        }

        static async System.Threading.Tasks.Task Main(string[] args)
        {
            LogMessage($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | CronJob starts");
            string siteUrl = "https://ddcosa.sharepoint.com/sites/EPMO2";
            //string listName = "Sample interface mgnt";
            string listName = "Interface Management";

            string username = "pol_admin@ddco.sa";
            string password = "Welcome@123$";
            string odataEndpoint = "https://ddcosa.sharepoint.com/sites/EPMO2/_api/ProjectData/"; // Replace with your OData URI
            string odataEndpoint1 = "https://ddcosa.sharepoint.com/sites/EPMO2/_api/ProjectServer/"; // Replace with your OData URI

            string odataUri = "https://ddcosa.sharepoint.com/sites/EPMO2/_api/ProjectData/Tasks";
            List<TaskDetails> tasks = new List<TaskDetails>();
            List<PrjDetails> prjs = new List<PrjDetails>();
            try
            {
                using (var context = new ClientContext(siteUrl))
                {
                    context.Credentials = new SharePointOnlineCredentials(username, new NetworkCredential("", password).SecurePassword);
                    string authCookie = (context.Credentials as SharePointOnlineCredentials)
                                        .GetAuthenticationCookie(new Uri(siteUrl));
                    if (string.IsNullOrEmpty(authCookie))
                    {
                        LogMessage("❌ Authentication failed. Cookie is empty.");
                        return;
                    }

                    LogMessage("✅ Successfully authenticated! Cookie retrieved.");

                    using (var handler = new HttpClientHandler { CookieContainer = new CookieContainer() })
                    {
                        handler.CookieContainer.SetCookies(new Uri(siteUrl), authCookie);
                        using (var httpClient = new HttpClient(handler))
                        {
                            string digestValue = await GetRequestDigest(httpClient, siteUrl);
                            if (string.IsNullOrEmpty(digestValue))
                            {
                                LogMessage("❌ Failed to retrieve X-RequestDigest.");
                                return;
                            }

                            LogMessage($"✅ X-RequestDigest retrieved: {digestValue}");
                            var result = await FetchAllInterfaceManagementItems(httpClient, siteUrl, listName);
                            /*List<string> taskGuids = result.Item1;
                            List<string> projGuids = result.Item2;

                            LogMessage("🎯 Task GUIDs Retrieved:" + taskGuids.Count());
                            List<TaskDetails> tasks = await FetchProjectTasks(httpClient, siteUrl, taskGuids);
                            foreach (var task1 in tasks)
                            {
                                Console.WriteLine(task1); // Calls overridden ToString() automatically
                                LogMessage(task1.ToString());
                            }
                            List<PrjDetails> prjs = await FetchDProject(httpClient, siteUrl, projGuids);

                            foreach (var prj1 in prjs)
                            {
                                Console.WriteLine(prj1); // Calls overridden ToString() automatically
                                LogMessage(prj1.ToString());
                            }*/

                            List<string> taskGuids = result.Item1;
                            List<string> projGuids = result.Item2;

                            LogMessage("🎯 Task GUIDs Retrieved:" + taskGuids.Count());

                            tasks = await FetchProjectTasks(httpClient, siteUrl, taskGuids);
                            prjs = await FetchDProject(httpClient, siteUrl, projGuids);
                            var lgmes = "";
                            try
                            {
                                // Load the SharePoint list
                                List list = context.Web.Lists.GetByTitle(listName);
                                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                                ListItemCollection listItems = list.GetItems(query);
                                context.Load(listItems);
                                context.ExecuteQuery();
                                TimeZoneInfo riyadhZone = TimeZoneInfo.FindSystemTimeZoneById("Arab Standard Time");
                                //TimeZoneInfo ksaTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Arab Standard Time");

                                foreach (ListItem listItem in listItems)
                                {
                                    var updateitm = 0;
                                    var updatetitle = 0;

                                    string taskGuid = listItem["Task_x0020_Name"]?.ToString();
                                    string projGuid = listItem["projectslist"]?.ToString();
                                    string depcflag = listItem["Dependency_x0020_Calculated_x002"]?.ToString();
                                    string depdate = listItem["Deliverable_x0020_Start_x0020_Da"]?.ToString();
                                    string dprojGuid = listItem["Project_x0020_Depends_x0020_on"]?.ToString();
                                    string dtaskGuid = listItem["Dependent_x0020_Task_x0020_Name"]?.ToString();

                                    lgmes += $"Source Project name: {listItem["Project_x0020_Name"]?.ToString()}\r\n" + Environment.NewLine;

                                    lgmes += $"Dependant Project name: {listItem["Project_x0020_Name_x0020_Depends"]?.ToString()}\r\n" + Environment.NewLine;


                                    if (listItem["Unique_x0020_ID"] == null || string.IsNullOrWhiteSpace(listItem["Unique_x0020_ID"].ToString()))
                                    {
                                        lgmes += $"Unique ID is empty\r\n" + Environment.NewLine;

                                        continue; // Skip to the next item
                                    }

                                    lgmes += "Unique ID - " + listItem["Unique_x0020_ID"] + "\r\n";
                                    lgmes += "--------------------------------------------------------------------\r\n";

                                    // 🔹 Compare source task
                                    if (!string.IsNullOrEmpty(taskGuid))
                                    {
                                        var taskData = tasks.FirstOrDefault(t => t.GUID == taskGuid);
                                        if (taskData != null)
                                        {
                                            lgmes += "Source Task Name:" + taskData.Title;
                                            var lstsrctitle = listItem["Project_x0020_Task_x0020_Name"].ToString();
                                            if (taskData.Title != lstsrctitle)
                                            {
                                                listItem["Project_x0020_Task_x0020_Name"] = taskData.Title;
                                                updatetitle = 1;
                                            }


                                            // Planned Date
                                            string taskStartDate = taskData.TStartDate?.ToString() ?? "N/A";
                                            if (taskStartDate != "N/A")
                                            {

                                                DateTime StaskStart = DateTime.Parse(taskStartDate).ToUniversalTime();
                                                DateTime riyadhTaskStart = TimeZoneInfo.ConvertTimeFromUtc(StaskStart, riyadhZone).Date;
                                                DateTime? ItemStaskStart = listItem["Planned_x0020_Date"] != null
    ? (DateTime?)TimeZoneInfo.ConvertTimeFromUtc(Convert.ToDateTime(listItem["Planned_x0020_Date"]), riyadhZone).Date
    : (DateTime?)null;

                                                // DateTime.Parse(listItem["Planned_x0020_Date"].ToString()).Date;
                                                /* DateTime? ItemStaskStart = listItem["Planned_x0020_Date"] != null
     ? Convert.ToDateTime(listItem["Planned_x0020_Date"]).Date
     : (DateTime?)null;*/
                                                DateTime odStaskStart = riyadhTaskStart.Date;
                                                lgmes += "StaskStart " + ItemStaskStart.ToString() + "compare with " + odStaskStart.ToString() + "\r\n" + Environment.NewLine;
                                                if (ItemStaskStart.ToString() != odStaskStart.ToString())
                                                {
                                                    //  Console.WriteLine("StaskStart " + ItemStaskStart.ToString() + "==" + odStaskStart.ToString());
                                                    lgmes += "StaskStart " + ItemStaskStart.ToString() + "==" + odStaskStart.ToString() + "\r\n" + Environment.NewLine;
                                                    listItem["Planned_x0020_Date"] = StaskStart;
                                                    updateitm = 1;
                                                }

                                            }

                                            string taskFinishDate = taskData.TFinishDate?.ToString() ?? "N/A";
                                            if (taskFinishDate != "N/A")
                                            {
                                                DateTime utcTaskFinish = DateTime.Parse(taskFinishDate).ToUniversalTime(); ;
                                                DateTime StaskFinish = TimeZoneInfo.ConvertTimeFromUtc(utcTaskFinish, riyadhZone).Date;


                                                //  DateTime StaskFinish = DateTime.Parse(taskFinishDate);
                                                //DateTime.Parse(listItem["Actual_x0020_Date"].ToString()).Date;
                                                /* DateTime? ItemStaskFinish = listItem["Actual_x0020_Date"] != null
     ? Convert.ToDateTime(listItem["Actual_x0020_Date"]).Date
     : (DateTime?)null;*/
                                                DateTime? ItemStaskFinish = listItem["Actual_x0020_Date"] != null
     ? (DateTime?)TimeZoneInfo.ConvertTimeFromUtc(Convert.ToDateTime(listItem["Actual_x0020_Date"]), riyadhZone).Date
     : (DateTime?)null;


                                                DateTime odStaskFinish = StaskFinish.Date;
                                                lgmes += "StaskFinish " + ItemStaskFinish.ToString() + "compare with " + odStaskFinish.ToString() + "\r\n" + Environment.NewLine;

                                                if (ItemStaskFinish.ToString() != odStaskFinish.ToString())
                                                {
                                                    //Console.WriteLine("StaskFinish" + ItemStaskFinish.ToString() + "==" + odStaskFinish.ToString());
                                                    lgmes += "StaskFinish" + ItemStaskFinish.ToString() + "==" + odStaskFinish.ToString() + "\r\n" + Environment.NewLine;

                                                    listItem["Actual_x0020_Date"] = StaskFinish;
                                                    updateitm = 1;
                                                }
                                            }





                                            string taskFinishPlusone = taskData.FDatePlusonee?.ToString() ?? "N/A";
                                            if (taskFinishPlusone != "N/A")
                                            {
                                                DateTime utcSfinplusOne = DateTime.Parse(taskFinishPlusone).ToUniversalTime(); ;
                                                // DateTime utcSfinplusOne = DateTime.Parse(taskFinishPlusone);
                                                DateTime SfinplusOne = TimeZoneInfo.ConvertTimeFromUtc(utcSfinplusOne, riyadhZone).Date;

                                                lgmes += "SfinplusOne" + SfinplusOne.ToString() + "\r\n";
                                                // lgmes += "ItemSfinplusOne" + listItem["Deliverable_x0020_Start_x0020_Da"].ToString() + "\r\n";
                                                /*  DateTime? ItemSfinplusOne = listItem["Deliverable_x0020_Start_x0020_Da"] != null
      ? Convert.ToDateTime(listItem["Deliverable_x0020_Start_x0020_Da"]).Date
      : (DateTime?)null;*/
                                                DateTime? ItemSfinplusOne = listItem["Deliverable_x0020_Start_x0020_Da"] != null
      ? (DateTime?)TimeZoneInfo.ConvertTimeFromUtc(Convert.ToDateTime(listItem["Deliverable_x0020_Start_x0020_Da"]), riyadhZone).Date
      : (DateTime?)null;
                                                lgmes += "ItemSfinplusOne" + ItemSfinplusOne.ToString() + "\r\n";

                                                DateTime odSfinplusOne = SfinplusOne.Date;
                                                lgmes += "StaskFinishPlusone " + ItemSfinplusOne.ToString() + "compare with " + odSfinplusOne.ToString() + "\r\n" + Environment.NewLine;

                                                if (depcflag != "Custom")
                                                {
                                                    if (ItemSfinplusOne.ToString() != odSfinplusOne.ToString())
                                                    {
                                                        //Console.WriteLine("SfinplusOne: " + ItemSfinplusOne.ToString() + "==" + odSfinplusOne.ToString());
                                                        lgmes += "SfinplusOne: " + ItemSfinplusOne.ToString() + "==" + odSfinplusOne.ToString() + "\r\n" + Environment.NewLine;
                                                        listItem["Deliverable_x0020_Start_x0020_Da"] = SfinplusOne;
                                                        updateitm = 1;
                                                    }
                                                }
                                            }







                                        }
                                    }

                                    if (!string.IsNullOrEmpty(dtaskGuid))
                                    {
                                        var dtaskData = tasks.FirstOrDefault(t => t.GUID == dtaskGuid);
                                        if (dtaskData != null)
                                        {
                                            // Extract task details (adjust fields based on response structure)
                                            string dtaskName = dtaskData.Title;
                                            lgmes += "DTaskname: " + dtaskName + " \r\n" + Environment.NewLine;
                                            var lstdpttitle = listItem["Task_x0020_Name_x0020_Depends_x0"].ToString();
                                            if (dtaskData.Title != lstdpttitle)
                                            {
                                                listItem["Task_x0020_Name_x0020_Depends_x0"] = dtaskData.Title;
                                                updatetitle = 1;
                                            }
                                            string dtaskStartDate = dtaskData.TStartDate?.ToString() ?? "N/A";
                                            if (dtaskStartDate != "N/A")
                                            {
                                                // DateTime DtaskStart = DateTime.Parse(dtaskStartDate);
                                                DateTime utcDtaskStart = DateTime.Parse(dtaskStartDate).ToUniversalTime(); ;

                                                DateTime DtaskStart = TimeZoneInfo.ConvertTimeFromUtc(utcDtaskStart, TimeZoneInfo.FindSystemTimeZoneById("Arab Standard Time")).Date;

                                                /*DateTime? ItemDtaskStart = listItem["Dependent_x0020_Task_x0020_Start"] != null
     ? Convert.ToDateTime(listItem["Dependent_x0020_Task_x0020_Start"]).Date
     : (DateTime?)null;*/
                                                DateTime? ItemDtaskStart = listItem["Dependent_x0020_Task_x0020_Start"] != null
                                                    ? (DateTime?)TimeZoneInfo.ConvertTimeFromUtc(Convert.ToDateTime(listItem["Dependent_x0020_Task_x0020_Start"]), TimeZoneInfo.FindSystemTimeZoneById("Arab Standard Time")).Date
                                                    : (DateTime?)null;


                                                DateTime odDtaskStart = DtaskStart;
                                                lgmes += "DtaskStart " + ItemDtaskStart.ToString() + "compare with " + odDtaskStart.ToString() + "\r\n" + Environment.NewLine;

                                                if (ItemDtaskStart.ToString() != odDtaskStart.ToString())
                                                {
                                                    //Console.WriteLine("DtaskStart" + ItemDtaskStart.ToString() + "==" + odDtaskStart.ToString());
                                                    lgmes += "DtaskStart" + ItemDtaskStart.ToString() + "==" + odDtaskStart.ToString() + "\r\n" + Environment.NewLine;

                                                    listItem["Dependent_x0020_Task_x0020_Start"] = DtaskStart;
                                                    updateitm = 1;
                                                }
                                            }
                                            string dtaskFinishDate = dtaskData.TFinishDate?.ToString() ?? "N/A";
                                            if (dtaskFinishDate != "N/A")
                                            {
                                                // DateTime DtaskFinish = DateTime.Parse(dtaskFinishDate);
                                                DateTime utcDtaskFinish = DateTime.Parse(dtaskFinishDate).ToUniversalTime(); ;
                                                DateTime DtaskFinish = TimeZoneInfo.ConvertTimeFromUtc(utcDtaskFinish, TimeZoneInfo.FindSystemTimeZoneById("Arab Standard Time")).Date;

                                                /* DateTime? ItemDtaskFinish = listItem["Dependent_x0020_Task_x0020_Finis"] != null
       ? Convert.ToDateTime(listItem["Dependent_x0020_Task_x0020_Finis"]).Date
       : (DateTime?)null; */
                                                DateTime? ItemDtaskFinish = listItem["Dependent_x0020_Task_x0020_Finis"] != null
    ? (DateTime?)TimeZoneInfo.ConvertTimeFromUtc(Convert.ToDateTime(listItem["Dependent_x0020_Task_x0020_Finis"]), TimeZoneInfo.FindSystemTimeZoneById("Arab Standard Time")).Date
    : (DateTime?)null;

                                                DateTime odDtaskFinish = DtaskFinish;
                                                lgmes += "DtaskFinish " + ItemDtaskFinish.ToString() + "compare with " + odDtaskFinish.ToString() + "\r\n" + Environment.NewLine;

                                                if (ItemDtaskFinish.ToString() != odDtaskFinish.ToString())
                                                {
                                                    Console.WriteLine("DtaskFinish " + ItemDtaskFinish.ToString() + "==" + odDtaskFinish);
                                                    lgmes += "DtaskFinish " + ItemDtaskFinish.ToString() + "==" + odDtaskFinish + "\r\n" + Environment.NewLine;

                                                    listItem["Dependent_x0020_Task_x0020_Finis"] = DtaskFinish;
                                                    updateitm = 1;
                                                }
                                            }


                                            Console.WriteLine($"Dependant Project Task Name: {dtaskName}");
                                        }
                                        else
                                        {
                                            Console.WriteLine("No 'd' object found in the response.");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Dependant Task GUID is null or empty.");
                                    }

                                    bool workflowPassed = false;
                                    if (!string.IsNullOrEmpty(dprojGuid))
                                    {
                                        var dproject = prjs.FirstOrDefault(p => p.GUID == dprojGuid);
                                        lgmes += "dep proj wf vars" + dproject.wfInP + dproject.PubByWf;
                                        if (dproject != null && dproject.wfInP == "Yes" && dproject.PubByWf == "Yes")
                                        {
                                            workflowPassed = true;
                                            lgmes += $"✅ Conditions of wf met for project: {dprojGuid}\r\n" + Environment.NewLine;
                                        }
                                        else
                                        {
                                            lgmes += $"❌ Conditions of wf NOT met for project: {dprojGuid}\r\n" + Environment.NewLine;
                                        }

                                    }
                                    if (workflowPassed) updateitm = 0;













                                    if (updateitm == 1)
                                    {
                                        Console.WriteLine("📝 Updating List Item...");
                                        lgmes += "Listitem got Updated" + Environment.NewLine;
                                        listItem["Approval_x0020_Email_x0020_Sent"] = "No";

                                     //   listItem.Update();

                                    }
                                    if(updatetitle == 1 && updateitm ==0)
                                    {
                                        Console.WriteLine("📝 Updating List Item...");
                                        lgmes += "Listitem title got Updated" + Environment.NewLine;

                                        //   listItem.Update();

                                    }
                                }
                                context.ExecuteQuery();

                                Console.WriteLine("All items updated successfully.");
                                LogMessage("All items updated successfully.");
                                lgmes += "All items updated successfully." + Environment.NewLine;

                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error: " + ex.Message);
                                LogMessage("Error 19: " + ex.Message);
                            }




                            lgmes += $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | CronJob Ends" + Environment.NewLine;

                            LogMessage(lgmes);
                            Thread.Sleep(5000);

                            Environment.Exit(0);

                        }

                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Error: {ex.Message}");
            }


        }
    }
}
