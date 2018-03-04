using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using System.IO;
using Excel=Microsoft.Office.Interop.Excel;
using Xltrail.Client.Models;
using System.Security.Cryptography;

namespace Xltrail.Client
{
    public static class Addin
    {
        [ComVisible(true)]
        public class RibbonController : ExcelRibbon, IExcelAddIn
        {
            static Excel.Application xlApp;
            static Dictionary<string, MenuItem> ids;
            static string XltrailPath = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), "xltrail");
            static string Staging = Path.Combine(XltrailPath, "staging");
            static string Workbooks = Path.Combine(XltrailPath, "workbooks");
            static string RepositoriesPath = Path.Combine(XltrailPath, "data");
            static string Repositories = Path.Combine(XltrailPath, "workbooks", "config.yaml");
            static string UserName = Environment.UserName;


            private Models.Config.Config Config;
            private List<Repository> GitRepositories;

            IRibbonUI ribbon;

            public void AutoOpen()
            {
                xlApp = (Excel.Application)ExcelDnaUtil.Application;
                xlApp.WorkbookActivate += XlApp_WorkbookActivate;
                ids = new Dictionary<string, MenuItem>();

                //load config.yaml
                var yaml = File.ReadAllText(Path.Combine(XltrailPath, "config.yaml"));
                var deserializer = new DeserializerBuilder().WithNamingConvention(new CamelCaseNamingConvention()).Build();
                Config = deserializer.Deserialize<Models.Config.Config>(yaml);

                Refresh();
            }


            public void AutoClose()
            {
                //backgroundThread.Abort();
            }


            public void Pull(string url, string path)
            {
                var credentials = Config.Credentials.Where(c => url.StartsWith(c.Url)).FirstOrDefault();
                if (!Directory.Exists(path))
                {
                    var cloneOptions = new LibGit2Sharp.CloneOptions();
                    if (credentials != null)
                    {
                        cloneOptions.CredentialsProvider = (_url, _user, _cred) => new LibGit2Sharp.UsernamePasswordCredentials
                        {
                            Username = credentials.Username,
                            Password = credentials.Password
                        };
                    }
                    LibGit2Sharp.Repository.Clone(url, path, cloneOptions);
                }
                else
                {
                    var fetchOptions = new LibGit2Sharp.FetchOptions();
                    if (credentials != null)
                    {
                        fetchOptions.CredentialsProvider = (_url, _user, _cred) => new LibGit2Sharp.UsernamePasswordCredentials
                        {
                            Username = credentials.Username,
                            Password = credentials.Password
                        };
                    }
                    using (var repository = new LibGit2Sharp.Repository(path))
                    {
                        foreach (var remote in repository.Network.Remotes)
                        {
                            IEnumerable<string> refSpecs = remote.FetchRefSpecs.Select(x => x.Specification);
                            LibGit2Sharp.Commands.Fetch(repository, remote.Name, refSpecs, fetchOptions, "");
                        }
                        foreach (var branch in repository.Branches.Where(b => b.IsRemote))
                        {
                            var localBranchName = branch.FriendlyName.Replace(branch.RemoteName + "/", "");
                            if (repository.Branches[localBranchName] == null)
                            {
                                repository.Branches.Update(
                                    repository.Branches.Add(localBranchName, branch.Tip),
                                    b => b.TrackedBranch = branch.CanonicalName);
                            }
                            var localBranch = repository.Branches[localBranchName];
                            if(localBranch.Tip != branch.Tip)
                            {
                                LibGit2Sharp.Commands.Checkout(repository, localBranchName);
                                repository.Reset(LibGit2Sharp.ResetMode.Hard, branch.Tip);
                            }
                        }

                    }
                }
            }

            public void PullRepository(Models.Config.Workbooks.Repository repository)
            {
                var path = Path.Combine(RepositoriesPath, repository.Alias);
                if (Directory.Exists(path))
                {
                    using (var repo = new LibGit2Sharp.Repository(path))
                    {
                        foreach (var remote in repo.Network.Remotes)
                        {
                            IEnumerable<string> refSpecs = remote.FetchRefSpecs.Select(x => x.Specification);
                            LibGit2Sharp.Commands.Fetch(repo, remote.Name, refSpecs, null, "");
                        }
                        foreach (var branch in repo.Branches.Where(b => b.IsRemote))
                        {
                            var localBranchName = branch.FriendlyName.Replace(branch.RemoteName + "/", "");
                            if (repo.Branches[localBranchName] == null)
                            {
                                repo.Branches.Update(
                                    repo.Branches.Add(localBranchName, branch.Tip),
                                    b => b.TrackedBranch = branch.CanonicalName);
                            }
                            LibGit2Sharp.Commands.Checkout(repo, localBranchName);
                            var signature = new LibGit2Sharp.Signature(UserName, UserName, new DateTimeOffset());
                            var mergeOptions = new LibGit2Sharp.MergeOptions();
                            repo.Merge(branch, signature, mergeOptions);
                        }

                    }
                }
                else
                {
                    var cloneOptions = new LibGit2Sharp.CloneOptions();
                    LibGit2Sharp.Repository.Clone(repository.Url, path, cloneOptions);
                }

            }


            public Models.Config.Workbooks RepositoriesConfig()
            {
                var yaml = File.ReadAllText(Repositories);
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(new CamelCaseNamingConvention())
                    .Build();
                return deserializer.Deserialize<Models.Config.Workbooks>(yaml);
            }

            public void PullRepositories()
            {
                var config = RepositoriesConfig();
                foreach (var repository in config.Repositories)
                    PullRepository(repository);
            }


            public void Synchronise()
            {
                var yaml = File.ReadAllText(Repositories);
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(new CamelCaseNamingConvention())
                    .Build();
                var workbooks = deserializer.Deserialize<Models.Config.Workbooks>(yaml);

                GitRepositories = new List<Repository>();
                var configuredRepositories = workbooks.Repositories.Select(r => r.Alias);

                foreach (var path in Directory.GetDirectories(RepositoriesPath))
                {
                    var repoName = Path.GetFileName(path);
                    if (configuredRepositories.Contains(repoName) && Repository.IsValid(path))
                    {
                        GitRepositories.Add(new Repository(path));
                    }
                }
            }

            public void Ribbon_Load(IRibbonUI ribbon)
            {
                this.ribbon = ribbon;
            }

            private void XlApp_WorkbookActivate(Excel.Workbook Wb)
            {
                ribbon.Invalidate();
            }

            public override string GetCustomUI(string RibbonID)
            {
                var id = CreateId();
                ids[id] = new MenuItem();
                var str = "<customUI onLoad='Ribbon_Load' xmlns='http://schemas.microsoft.com/office/2006/01/customui'>\n";
                str += "<ribbon>\n";
                str += "<tabs>\n";
                str += "<tab id='tab' label='Xltrail'>\n";
                str += "<group id='group1' label='Workbooks'>\n";
                str += "<dynamicMenu id='" + id + "' label='Workbooks' imageMso='MicrosoftExcel' size='large' getContent='BuildMenu' />\n";
                str += "</group>";
                str += "<group id='group2' label='Save' getVisible='GetWorkbookVisibility'>\n";
                str += "<button id='workbookName' getLabel='GetActiveWorkbookName' size='normal' imageMso='Info' />\n";
                str += "<button id='commitButton' label='Commit' size='normal' imageMso='FileSave' onAction='CommitButton_Click' />\n";
                str += "</group>";
                str += "</tab>";
                str += "</tabs>";
                str += "</ribbon>";
                str += "</customUI>";
                return str;
            }

            private string CreateId()
            {
                var id = Convert.ToBase64String(Guid.NewGuid().ToByteArray());
                id = "id-" + id.Replace("+", "-").Replace("/", "_").Replace(" ", "").Replace("=", "");
                return id;
            }

            private string StagedWorkbookPath(string repository, string branch, string workbookPath)
            {
                var path = Path.Combine(
                    Environment.GetEnvironmentVariable("LocalAppData"),
                    "xltrail",
                    "staging",
                    repository,
                    workbookPath);

                return Path.Combine(path, Path.GetFileNameWithoutExtension(workbookPath)
                    + "_" + branch.Replace("origin/", "")
                    + Path.GetExtension(workbookPath));
            }

            private string GetSHA1Hash(Stream stream)
            {
                using (SHA1Managed sha = new SHA1Managed())
                {
                    byte[] checksum = sha.ComputeHash(stream);
                    return BitConverter.ToString(checksum)
                        .Replace("-", string.Empty);
                }
            }

            private string GetSHA1Hash(string filename)
            {
                using (FileStream stream = File.OpenRead(filename))
                {
                    return GetSHA1Hash(stream);
                }
            }

            private bool IsWorkbook(string path)
            {
                return (
                    path.EndsWith(".xls") ||
                    path.EndsWith(".xlsb") ||
                    path.EndsWith(".xlsm") ||
                    path.EndsWith(".xlsm") ||
                    path.EndsWith(".xla") ||
                    path.EndsWith(".xlam"));
            }

            public string BuildMenu(IRibbonControl control)
            {
                ribbon.Invalidate();
                var menuItem = ids[control.Id];
                
                var str = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui'>";
                if (menuItem.IsRoot())
                {
                    string id = null;
                    foreach(var repository in GitRepositories)
                    {
                        id = CreateId();
                        ids[id] = new MenuItem(repository);
                        str += "<dynamicMenu id='" + id + "' label='" + Path.GetFileName(repository.Path) + "' imageMso='Folder' getContent='BuildMenu' />\n";
                    }
                    id = CreateId();
                    ids[id] = new MenuItem();
                    if(GitRepositories.Count > 0)
                        str += "<menuSeparator id='separator' />";
                     str += "<button id='" + id + "' label='Refresh' imageMso='Repeat' onAction='Refresh_Click' />\n";
                }
                else if (menuItem.IsWorkbook())
                {
                    var repository = menuItem.Repository;
                    var workbook = menuItem.Workbook;

                    foreach (var branch in menuItem.Workbook.Branches)
                    {
                        var id = CreateId();
                        var fileName = Path.GetFileName(workbook.Path);
                        ids[id] = new MenuItem(menuItem.Repository, new Workbook(menuItem.Workbook.Path, branch));
                        str += "<button id='" + id + "' label='" + branch + "' imageMso='MicrosoftExcel' onAction='OpenWorkbook_Click' />\n";
                    }
                }
                else
                {
                    //IsFolder() == True
                    var repository = menuItem.Repository;
                    var folder = repository.Workbooks(menuItem.Folder);

                    foreach (var f in folder.Folders.OrderBy(x => x))
                    {
                        var id = CreateId();
                        ids[id] = new MenuItem(repository, f);
                        var fileName = Path.GetFileName(f);
                        str += "<dynamicMenu id='" + id + "' label='" + fileName + "' imageMso='Folder' getContent='BuildMenu' />\n";
                    }

                    foreach (var workbook in folder.Workbooks.OrderBy(w => w.Path))
                    {
                        var id = CreateId();
                        var fileName = Path.GetFileName(workbook.Path);
                        ids[id] = new MenuItem(menuItem.Repository, workbook);
                        if (workbook.Branches.Count == 1)
                        {
                            str += "<button id='" + id + "' label='" + fileName + "' imageMso='MicrosoftExcel' onAction='OpenWorkbook_Click' />\n";
                        }
                        else
                        {
                            str += "<dynamicMenu id='" + id + "' label='" + fileName + "' imageMso='MicrosoftExcel' getContent='BuildMenu' />\n";
                        }
                    }
                }
                str += "</menu>";
                return str;
            }

            public string GetActiveWorkbookName(IRibbonControl control)
            {
                var path = xlApp.ActiveWorkbook.FullName;
                if (!path.Contains(Staging))
                    return "(not a git workbook)";
                var fileName = Path.GetFileNameWithoutExtension(path);
                var fileExtension = Path.GetExtension(path);
                var parts = fileName.Split('_');
                var branch = parts.Last();
                return fileName.Substring(0, fileName.Length - branch.Length - 1) + fileExtension + " (" + branch + ")";
            }

            public bool GetWorkbookVisibility(IRibbonControl control)
            {
                var path = xlApp.ActiveWorkbook.FullName;
                if (!path.Contains(Staging))
                    return false;
                return true;
            }

            public void Refresh_Click(IRibbonControl control)
            {
                var cursor = xlApp.Cursor;
                try
                {
                    xlApp.Cursor = Excel.XlMousePointer.xlWait;
                    Refresh();
                }
                catch (Exception ex)
                {
                    xlApp.Cursor = cursor;
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
                finally
                {
                    xlApp.Cursor = cursor;
                }
            }


            public void CommitButton_Click(IRibbonControl control)
            {
                var cursor = xlApp.Cursor;
                try
                {
                    xlApp.Cursor = Excel.XlMousePointer.xlWait;
                    CommitAndPushWorkbook(xlApp.ActiveWorkbook);
                }
                catch(Exception ex)
                {
                    xlApp.Cursor = cursor;
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
                finally
                {
                    xlApp.Cursor = cursor;
                }
            }

            public void OpenWorkbook_Click(IRibbonControl control)
            {
                var cursor = xlApp.Cursor;
                try
                {
                    xlApp.Cursor = Excel.XlMousePointer.xlWait;
                    var menuItem = ids[control.Id];
                    var workbook = OpenWorkbook(menuItem);
                    xlApp.Workbooks.Open(workbook);
                }
                catch (Exception ex)
                {
                    xlApp.Cursor = cursor;
                    System.Windows.Forms.MessageBox.Show(ex.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
                finally
                {
                    xlApp.Cursor = cursor;
                }
            }


            public string OpenWorkbook(MenuItem menuItem)
            {
                //path to staged workbook file
                var path = Path.Combine(
                    Staging,
                    menuItem.Repository.Path.Split('\\').Last(),
                    menuItem.Workbook.Path);

                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                var fileName = Path.Combine(path, Path.GetFileNameWithoutExtension(menuItem.Workbook.Path)
                    + "_" + menuItem.Workbook.Branches.First().Replace("origin/", "")
                    + Path.GetExtension(menuItem.Workbook.Path));

                if (!File.Exists(fileName))
                {
                    var repository = menuItem.Repository;
                    var workbook = menuItem.Workbook;
                    var branch = menuItem.Workbook.Branches.First();

                    //branch could be "branch" or "origin/branch"
                    var branches = repository.GitRepository.Branches.Select(b => b.FriendlyName);
                    if (!branches.Contains(branch))
                        branch = "origin/" + branch;
                    var treeEntry = repository.GitRepository.Branches[branch][workbook.Path];
                    var blob = (LibGit2Sharp.Blob)treeEntry.Target;
                    var contentStream = blob.GetContentStream();
                    using (var fileStream = File.Create(fileName))
                    {
                        contentStream.Seek(0, SeekOrigin.Begin);
                        contentStream.CopyTo(fileStream);
                    }
                }
                return fileName;
            }

            public void CommitAndPushWorkbook(Excel.Workbook workbook)
            {
                //refresh repository definitions
                var config = RepositoriesConfig();

                //save file
                workbook.Save();

                //get reponame, file path and branch name
                var path = workbook.FullName;
                var fileName = Path.GetFileNameWithoutExtension(path);
                var fileExtension = Path.GetExtension(path);
                var parts = fileName.Split('_');
                var branchName = parts.Last();

                fileName = fileName.Substring(0, fileName.Length - branchName.Length - 1) + fileExtension;
                var repositoryName = path.Substring(Staging.Length+1).Split('\\').First();
                var repository = config.Repositories.Where(r => r.Alias == repositoryName).FirstOrDefault();

                if (repository == null)
                    throw new Exception("Unknown repository: " + repositoryName);

                //update repository from remote to avoid conflicts
                PullRepository(repository);

                //workbook path inside repository
                var filePath = Path.GetDirectoryName(Path.GetDirectoryName(path)).Substring(Path.Combine(Staging, repositoryName).Length);
                
                var fileRepoPath = Path.Combine(RepositoriesPath, repositoryName);
                if(filePath.Length > 0)
                    fileRepoPath = Path.Combine(fileRepoPath, filePath);
                fileRepoPath = Path.Combine(fileRepoPath, fileName);

                //get repository
                var gitRepository = new LibGit2Sharp.Repository(Path.Combine(RepositoriesPath, repositoryName));
                LibGit2Sharp.Commands.Checkout(gitRepository, branchName);

                //pull from remote (to avoid conflicts)

                //copy file from staging => repository
                File.Copy(path, fileRepoPath, true);

                //stage
                LibGit2Sharp.Commands.Stage(gitRepository, fileRepoPath);

                //commit
                var author = new LibGit2Sharp.Signature(UserName, UserName, new DateTimeOffset());
                var commitOptions = new LibGit2Sharp.CommitOptions();
                var commitMessage = "Updated " + fileName;

                gitRepository.Commit(commitMessage, author, author, commitOptions);

                //get credentials
                var pushUrl = gitRepository.Network.Remotes["origin"].PushUrl;
                var credentials = Config.Credentials.Where(c => pushUrl.StartsWith(c.Url)).FirstOrDefault();

                LibGit2Sharp.PushOptions pushOptions = new LibGit2Sharp.PushOptions();
                if(credentials != null)
                {
                    pushOptions.CredentialsProvider = new LibGit2Sharp.Handlers.CredentialsHandler(
                        (url, usernameFromUrl, types) =>
                            new LibGit2Sharp.UsernamePasswordCredentials()
                            {
                                Username = credentials.Username,
                                Password = credentials.Password
                            });
                }
                gitRepository.Network.Push(gitRepository.Branches[branchName], pushOptions);
            }


            public void Refresh()
            {
                Pull(Config.Workbooks, Workbooks); //update config
                PullRepositories();
                Synchronise();
            }


        }
    }
}
