using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DevExpress.Export.Xl;
using Octokit;
using FileMode = System.IO.FileMode;

namespace GitHubAnalyzerTool {
    class Program {
        static string name = "repository_analyzer";
        static void Main(string[] args) {
            Console.WriteLine("configuration...");
            string text = File.ReadAllText("keywords.cnfg");
            string[] keywords = text.Split(',');
            Console.WriteLine("configuration complete");
            Console.WriteLine("search...");
            string document = "github_search_result.xlsx";
            IXlExporter exporter = XlExport.CreateExporter(XlDocumentFormat.Xlsx);
            using(FileStream fs = new FileStream(document, FileMode.Create, FileAccess.ReadWrite)) 
                CreateDocument(exporter, fs, keywords);
            Console.WriteLine("search complete");
            Console.WriteLine("press and key to run document...");
            Process.Start(document);
        }
        static void CreateDocument(IXlExporter exporter, FileStream fs, string[] keywords) {
            using(IXlDocument document = exporter.CreateDocument(fs)) {
                for(int i = 0; i < keywords.Length; i++)
                    CreateTermSheet(document, keywords[i], SearchRepositories(keywords[i]).Result);
            }
        }
        static void CreateTermSheet(IXlDocument document, string keyword, SearchRepositoryResult result) {
            using(IXlSheet sheet = document.CreateSheet()) {
                sheet.Name = keyword;
                CreateHeader(sheet);
                for(int r = 0; r < result.Items.Count; r++) {
                    var repo = result.Items[r];
                    if(repo.Owner.Type == AccountType.User) {
                        using(IXlRow row = sheet.CreateRow()) {
                            CreateCell(row, repo.Name);
                            CreateCell(row, repo.GitUrl);
                            CreateCell(row, repo.Owner.Login);
                            CreateCell(row, repo.Owner.Bio);
                            CreateCell(row, repo.Owner.Location);
                            CreateCell(row, repo.Owner.Company);
                            CreateCell(row, repo.Owner.Email);
                            CreateCell(row, repo.Owner.HtmlUrl);
                        }
                    }
                }
            }
        }
        static void CreateHeader(IXlSheet sheet) {
            using(IXlRow row = sheet.CreateRow()) {
                CreateHeaderCell(row, nameof(Repository.Name));
                CreateHeaderCell(row, nameof(Repository.Url));
                CreateHeaderCell(row, nameof(Account.Login));
                CreateHeaderCell(row, nameof(Account.Bio));
                CreateHeaderCell(row, nameof(Account.Location));
                CreateHeaderCell(row, nameof(Account.Company));
                CreateHeaderCell(row, nameof(Account.Email));
                CreateHeaderCell(row, nameof(Account.HtmlUrl));
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
            var github = new GitHubClient(new ProductHeaderValue(name));
            var request = new SearchRepositoriesRequest(term) {
                Language = Language.CSharp,
                SortField = RepoSearchSort.Stars,
                Order = SortDirection.Descending,
                Size = Range.GreaterThanOrEquals(100)
            };
            return await github.Search.SearchRepo(request);
        }
    }
}
