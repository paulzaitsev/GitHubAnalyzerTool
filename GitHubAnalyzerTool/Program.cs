using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DevExpress.Export.Xl;
using Octokit;
using FileMode = System.IO.FileMode;

namespace GitHubAnalyzerTool {
    class SearchResult {
        public Repository Repository { get; set; }
        public List<User> Contributors { get; set; } = new List<User>();
    }

    class Program {
        static string name = "repository_analyzer";
        static GitHubClient client;
        static List<SearchResult> searchResults;
        static Task[] tasks;
        static void Main(string[] args) {
            Console.WriteLine("configuration...");
            string text = File.ReadAllText("keywords.cnfg");
            string[] keywords = text.Split(',');
            string authToken = File.ReadAllText("auth");
            tasks = new Task[keywords.Length];
            searchResults = new List<SearchResult>(512);
            Console.WriteLine("configuration complete");
            Console.WriteLine("search started...");
            string docName = "github_search_result.xlsx";
            StartSearch(keywords, authToken);
            Task.WaitAll(tasks);
            Console.WriteLine("search completed");
            Console.WriteLine("exporting...");
            IXlExporter exporter = XlExport.CreateExporter(XlDocumentFormat.Xlsx);
            using(FileStream fs = new FileStream(docName, FileMode.Create, FileAccess.ReadWrite))
                Export(exporter, fs);
            Console.WriteLine("export completed");
            Console.WriteLine("press and key to open document...");
            Console.ReadKey();
            Process.Start(docName);
        }
        static void Export(IXlExporter exporter, FileStream fs) {
            using(IXlDocument document = exporter.CreateDocument(fs)) {
                using(IXlSheet sheet = document.CreateSheet()) {
                    CreateHeader(sheet);
                    for(var i = 0; i < searchResults.Count; i++) {
                        var item = searchResults[i];
                        using(IXlRow row = sheet.CreateRow())
                            WriteRepositoryData(row, item.Repository);
                        for(int index = 0; index < item.Contributors.Count; index++) {
                            using(IXlRow row = sheet.CreateRow()) {
                                var contributor = item.Contributors[index];
                                CreateCell(row, null);
                                CreateCell(row, null);
                                WriteUserData(row, contributor);
                            }
                        }
                    }
                    ApplyAutoFilter(sheet);
                }
            }
        }
        static void StartSearch(string[] keywords, string authToken) {
            client = new GitHubClient(new ProductHeaderValue(name));
            client.Credentials = new Credentials(authToken);
            for(int i = 0; i < keywords.Length; i++) {
                string keyword = keywords[i];
                var task = Task.Run(() => SearchRepositories(keyword));
                tasks[i] = task.ContinueWith(t => { FindContributors(t, keyword); });
            }
        }
        static void FindContributors(Task<SearchRepositoryResult> task, string keyword) {
            if(task.Status != TaskStatus.RanToCompletion)
                return;
            SearchRepositoryResult result = task.Result;
            if(result == null)
                return;
            Console.WriteLine($"search for keyword: '{keyword}'");
            for(int r = 0; r < result.Items.Count; r++) {
                var repository = result.Items[r];
                if(repository.Owner.Type != AccountType.Organization) {
                    SearchResult sr = new SearchResult();
                    sr.Repository = repository;
                    sr.Contributors = new List<User>();
                    var list = client.Repository.GetAllContributors(repository.Id);
                    var listResult = list.Result;
                    for(int i = 0; i < listResult.Count; i++) {
                        var userTask = client.User.Get(listResult[i].Login);
                        sr.Contributors.Add(userTask.Result);
                    }
                    searchResults.Add(sr);
                }
            }
            Console.WriteLine($"keyword: '{keyword}' processed");
        }
        static void WriteRepositoryData(IXlRow row, Repository repo) {
            CreateCell(row, repo.Name);
            CreateCell(row, repo.GitUrl);
        }
        static void WriteUserData(IXlRow row, User userProfile) {
            CreateCell(row, userProfile.Name);
            CreateCell(row, userProfile.HtmlUrl);
            CreateCell(row, userProfile.Location);
            CreateCell(row, userProfile.Company);
            CreateCell(row, IsOwnerHireable(userProfile));
            CreateCell(row, userProfile.Bio);
            CreateCell(row, userProfile.Email);
        }
        static void ApplyAutoFilter(IXlSheet sheet) {
            int right = sheet.CurrentColumnIndex - 1;
            int bottom = sheet.CurrentRowIndex - 1;
            sheet.AutoFilterRange = XlCellRange.FromLTRB(0, 0, right, bottom);
        }
        static string IsOwnerHireable(User user) {
            return user.Hireable.HasValue ? user.Hireable.Value.ToString() : "no info";
        }
        static void CreateHeader(IXlSheet sheet) {
            using(IXlRow row = sheet.CreateRow()) {
                CreateHeaderCell(row, nameof(Repository.Name));
                CreateHeaderCell(row, nameof(Repository.Url));

                CreateHeaderCell(row, nameof(Account.Name));
                CreateHeaderCell(row, nameof(Account.HtmlUrl));
                CreateHeaderCell(row, nameof(Account.Location));
                CreateHeaderCell(row, nameof(Account.Company));
                CreateHeaderCell(row, nameof(Account.Hireable));
                CreateHeaderCell(row, nameof(Account.Bio));
                CreateHeaderCell(row, nameof(Account.Email));
            }
        }
        static void CreateCell(IXlRow row, string val) {
            using(IXlCell cell = row.CreateCell())
                cell.Value = val;
        }
        static void CreateHeaderCell(IXlRow row, string val) {
            using(IXlCell cell = row.CreateCell()) {
                cell.Value = val;
                cell.Formatting = new XlCellFormatting();
                cell.Formatting.Font = new XlFont();
                cell.Formatting.Font.Bold = true;
            }
        }
        static async Task<SearchRepositoryResult> SearchRepositories(string term) {
            var request = new SearchRepositoriesRequest(term) {
                Language = Language.CSharp,
                SortField = RepoSearchSort.Stars,
                Order = SortDirection.Descending
            };
            return await client.Search.SearchRepo(request);
        }
    }
}
