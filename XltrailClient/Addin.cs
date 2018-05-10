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
using Xltrail.Client.Providers;
using System.Security.Cryptography;
using log4net;
using System.Threading;

namespace Xltrail.Client
{
    public static class Addin
    {
        [ComVisible(true)]
        public class RibbonController : ExcelRibbon, IExcelAddIn
        {
            static Excel.Application xlApp;
            static string XltrailPath = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), "xltrail");
            static string ConfigPath = Path.Combine(XltrailPath, "config.yaml");

            static string StagingPath = Path.Combine(XltrailPath, "staging");
            static string WorkbooksPath = Path.Combine(XltrailPath, "config");
            static string RepositoriesPath = Path.Combine(XltrailPath, "repositories");
            static string LogsPath = Path.Combine(XltrailPath, "logs");

            static string Repositories = Path.Combine(XltrailPath, "config", "config.yaml");

            private Models.Config.Config Config;
            private Repositories repositories;

            private static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

            IRibbonUI ribbon;


            private void Initialise()
            {
                if (!Directory.Exists(StagingPath))
                    Directory.CreateDirectory(StagingPath);

                if (!Directory.Exists(RepositoriesPath))
                    Directory.CreateDirectory(RepositoriesPath);

                if (!Directory.Exists(LogsPath))
                    Directory.CreateDirectory(LogsPath);
            }

            public void AutoOpen()
            {
                Initialise();
                Logger.Setup();
                logger.Info("Starting Addin");
                xlApp = (Excel.Application)ExcelDnaUtil.Application;
                xlApp.WorkbookActivate += XlApp_WorkbookActivate;
                xlApp.WorkbookAfterSave += XlApp_WorkbookAfterSave;
                Config = LoadConfig();
                RefreshAll();
            }

            private void XlApp_WorkbookAfterSave(Excel.Workbook Wb, bool Success)
            {
                var path = Wb.FullName;
                if (!path.StartsWith(StagingPath))
                    return;

                //get workbookBranch from workbook path/filename
                var workbookPath = path.Substring(StagingPath.Length + 1).Replace("\\", "/");
                var workbookBranch = repositories.GetWorkbookVersionFromPath(workbookPath);

                //circuit breaker
                if (workbookBranch == null)
                    return;

                //get credentials
                var pushUrl = workbookBranch.Workbook.Repository.GitRepository.Network.Remotes["origin"].PushUrl;
                var credentials = Config.Credentials.Where(c => pushUrl.StartsWith(c.Url)).FirstOrDefault();

                //commit
                workbookBranch.Commit(path, credentials.Username ?? Environment.UserName, credentials.Email);
            }

            private void ShowNotification(string description)
            {
                var notification = new System.Windows.Forms.NotifyIcon()
                {
                    Visible = true,
                    Icon = System.Drawing.SystemIcons.Information,
                    // optional - BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info,
                    // optional - BalloonTipTitle = "My Title",
                    BalloonTipText = description,
                };

                // Display for 5 seconds.
                notification.ShowBalloonTip(5000);

                // This will let the balloon close after it's 5 second timeout
                // for demonstration purposes. Comment this out to see what happens
                // when dispose is called while a balloon is still visible.
                Thread.Sleep(10000);

                // The notification should be disposed when you don't need it anymore,
                // but doing so will immediately close the balloon if it's visible.
                notification.Dispose();

            }


            public void AutoClose()
            {
                //backgroundThread.Abort();
            }


            public Models.Config.Config LoadConfig()
            {
                //load config.yaml
                logger.InfoFormat("Load config from {0}", ConfigPath);
                if (!File.Exists(ConfigPath))
                {
                    logger.InfoFormat("Config not found, use defaults");
                    return new Models.Config.Config();
                }

                var yaml = File.ReadAllText(ConfigPath);
                var deserializer = new DeserializerBuilder().WithNamingConvention(new CamelCaseNamingConvention()).Build();
                return deserializer.Deserialize<Models.Config.Config>(yaml);
            }


            public Models.Config.Workbooks LoadRepositoriesConfig()
            {
                logger.InfoFormat("Load repositores config from {0}", Repositories);
                var yaml = File.ReadAllText(Repositories);
                logger.Info(yaml);
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(new CamelCaseNamingConvention())
                    .Build();
                return deserializer.Deserialize<Models.Config.Workbooks>(yaml);
            }

            /// <summary>
            /// Synchronise config => list of repositories
            /// </summary>
            public void ReadRepositoriesFromFilesystem()
            {
                var yaml = File.ReadAllText(Repositories);
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(new CamelCaseNamingConvention())
                    .Build();
                var workbooks = deserializer.Deserialize<Models.Config.Workbooks>(yaml);

                //get list of configured repositories
                var configuredRepositories = workbooks.Repositories.Select(r => r.Alias);
                repositories = new Repositories(RepositoriesPath, configuredRepositories);
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
                var str = "<customUI onLoad='Ribbon_Load' xmlns='http://schemas.microsoft.com/office/2006/01/customui'>\n";
                str += "<ribbon>\n";
                str += "<tabs>\n";
                str += "<tab id='tab' label='Xltrail'>\n";
                str += "<group id='group1' label='Workbooks'>\n";
                str += "<dynamicMenu id='id-root' label='Workbooks' imageMso='MicrosoftExcel' size='large' getContent='BuildMenu' />\n";
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


            public string BuildWorkbookMenu(IRibbonControl control)
            {
                ribbon.Invalidate();
                var workbook = repositories.GetWorkbookFromId(control.Id);
                var str = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui'>\n";
                foreach (var branch in workbook.Branches.Where(b => b.IsStagingBranch || b.HasConflict))
                {
                    str += "<button id='" + branch.Id + "' label='" + branch.DisplayName + "' imageMso='MicrosoftExcel' onAction='OpenWorkbook_Click' />\n";
                }
                str += "</menu>";
                return str;
            }


            public string BuildMenu(IRibbonControl control)
            {
                ribbon.Invalidate();
                
                var str = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui'>";
                if (control.Id == "id-root")
                {
                    foreach(var repository in repositories)
                    {
                        str += "<dynamicMenu id='" + repository.Id + "' label='" + repository.Name + "' imageMso='Folder' getContent='BuildMenu' />\n";
                    }
                    if(repositories.Count() > 0)
                        str += "<menuSeparator id='separator' />";
                     str += "<button id='id-refresh' label='Refresh' imageMso='Repeat' onAction='Refresh_Click' />\n";
                }
                else
                {
                    var repositoryAndFolder = repositories.GetRepositoryFromId(control.Id);
                    var repository = repositoryAndFolder.Item1;
                    var path = repositoryAndFolder.Item2;

                    var folders = repository.GetFolders(path).OrderBy(x => x);
                    foreach (var f in folders)
                    {
                        var name = Path.GetFileName(f);
                        var id = repository.Folders[path + (path != "" ? "/" : "") + f];
                        str += "<dynamicMenu id='" + id + "' label='" + name + "' imageMso='Folder' getContent='BuildMenu' />\n";
                    }

                    foreach (var workbook in repository.GetWorkbooks(path).OrderBy(w => w.Path))
                    {
                        var fileName = Path.GetFileName(workbook.Path);
                        if (workbook.Branches.Count == 1)
                        {
                            str += "<button id='" + workbook.Branches.First().Id + "' label='" + fileName + "' imageMso='MicrosoftExcel' onAction='OpenWorkbook_Click' />\n";
                        }
                        else
                        {
                            str += "<dynamicMenu id='" + workbook.Id + "' label='" + fileName + "' imageMso='MicrosoftExcel' getContent='BuildWorkbookMenu' />\n";
                        }
                    }
                }
                str += "</menu>";
                return str;
            }


            public string GetActiveWorkbookName(IRibbonControl control)
            {
                var path = xlApp.ActiveWorkbook.FullName;
                if (!path.Contains(StagingPath))
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
                if (!path.Contains(StagingPath))
                    return false;
                return true;
            }

            public void Refresh_Click(IRibbonControl control)
            {
                var cursor = xlApp.Cursor;
                try
                {
                    xlApp.Cursor = Excel.XlMousePointer.xlWait;
                    RefreshAll();
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
                    var workbookVersion = repositories.GetWorkbookVersionFromId(control.Id);
                    var workbookPath = OpenWorkbook(workbookVersion);
                    if (!workbookVersion.IsStagingBranch)
                    {
                        xlApp.Workbooks.Open(workbookPath, false, true);
                    }
                    else
                    {
                        xlApp.Workbooks.Open(workbookPath);
                    }
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

            public string OpenWorkbook(Branch branch)
            {
                //path to staged workbook file
                var fileName = Path.Combine(StagingPath, branch.Workbook.Repository.Name, branch.Path);
                var dirName = Path.GetDirectoryName(fileName);

                if (!Directory.Exists(dirName))
                    Directory.CreateDirectory(dirName);

                if (!File.Exists(fileName))
                {
                    //copy file to staging area
                    var repository = branch.Workbook.Repository;
                    var workbook = branch.Workbook;
                    //var branch = menuItem.Workbook.Branches.First();

                    //get workbook's latest commit sha
                    var head = branch.Head;

                    //branch could be "branch" or "origin/branch"
                    /*
                    var branches = repository.GitRepository.Branches.Select(b => b.FriendlyName).ToList();
                    if (!branches.Contains(branch))
                        branch = "origin/" + branch;
                    */

                    //get blob and write to filesystem
                    var treeEntry = repository.GitRepository.Branches[branch.Name][workbook.Path];
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
                var config = LoadRepositoriesConfig();

                //save file
                workbook.Save();

                //get reponame, file path and branch name
                var path = workbook.FullName;
                var fileName = Path.GetFileNameWithoutExtension(path);
                var fileExtension = Path.GetExtension(path);
                var parts = fileName.Split('_');
                var origin = parts.Last();
                var branchName = parts[parts.Count() - 1];

                fileName = fileName.Substring(0, fileName.Length - branchName.Length - 1) + fileExtension;
                var repositoryName = path.Substring(StagingPath.Length+1).Split('\\').First();
                var repository = config.Repositories.Where(r => r.Alias == repositoryName).FirstOrDefault();

                if (repository == null)
                    throw new Exception("Unknown repository: " + repositoryName);

                //update repository from remote to avoid conflicts
                var repositoryPath = Path.Combine(RepositoriesPath, repository.Alias);
                var credentials = Config.Credentials.Where(c => repository.Url.StartsWith(c.Url)).FirstOrDefault();
                GitProvider.PullFromRemote(repository.Url, repositoryPath, credentials);


                //workbook path inside repository
                var filePath = Path.GetDirectoryName(Path.GetDirectoryName(path)).Substring(Path.Combine(StagingPath, repositoryName).Length);
                
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

                //get credentials
                var pushUrl = gitRepository.Network.Remotes["origin"].PushUrl;
                credentials = Config.Credentials.Where(c => pushUrl.StartsWith(c.Url)).FirstOrDefault();

                //commit
                var author = new LibGit2Sharp.Signature(credentials.Username ?? Environment.UserName, credentials.Email, DateTime.Now);
                var commitOptions = new LibGit2Sharp.CommitOptions();
                var commitMessage = "Updated " + fileName;
                var committer = author;

                gitRepository.Commit(commitMessage, author, committer, commitOptions);

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


            public void RefreshAll()
            {
                try
                {
                    logger.Info("Refresh config from remote");
                    if (Config.Repositories != null)
                    {
                        logger.InfoFormat("Update config from {0}", Config.Repositories);
                        GitProvider.PullFromRemote(
                            Config.Repositories,
                            WorkbooksPath,
                            Config.Credentials.Where(c => Config.Repositories.StartsWith(c.Url)).FirstOrDefault());
                    }
                    else
                    {
                        logger.Info("No Git remote repository defined for workbooks config");
                    }

                    var repositoriesConfig = LoadRepositoriesConfig();
                    foreach (var repository in repositoriesConfig.Repositories)
                    {
                        var repositoryPath = Path.Combine(RepositoriesPath, repository.Alias);
                        var credentials = Config.Credentials.Where(c => repository.Url.StartsWith(c.Url)).FirstOrDefault();
                        GitProvider.PullFromRemote(repository.Url, repositoryPath, credentials);
                        GitProvider.EnsureStagingBranches(repositoryPath);
                    }

                    ReadRepositoriesFromFilesystem();
                }
                catch(Exception ex)
                {
                    logger.Error(ex);
                }
            }


        }
    }
}
