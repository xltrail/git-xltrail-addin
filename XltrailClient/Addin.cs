using System;
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
using Config = Xltrail.Client.Models.Config;
using System.Security.Cryptography;
using log4net;
using System.Threading;
using System.Collections.Generic;

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
            static string ConfigPathRepositories = Path.Combine(XltrailPath, "config", "config.yaml");

            static string StagingPath = Path.Combine(XltrailPath, "staging");
            static string RepositoriesPath = Path.Combine(XltrailPath, "repositories");
            static string LogsPath = Path.Combine(XltrailPath, "logs");
            static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

            private Models.Config.Config Config;
            private Repositories repositories;
            private Branch activeWorkbookBranch;

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

            private string GetWorkbookPath(string path)
            {
                return path.Substring(StagingPath.Length + 1).Replace("\\", "/");
            }

            private void XlApp_WorkbookAfterSave(Excel.Workbook Wb, bool Success)
            {
                var path = Wb.FullName;
                if (!path.StartsWith(StagingPath))
                    return;

                //get workbookBranch from workbook path/filename
                var workbookPath = GetWorkbookPath(path);
                var workbookBranch = repositories.GetWorkbookVersionFromPath(workbookPath);

                //circuit breaker
                if (workbookBranch == null)
                    return;

                //get credentials
                var pushUrl = workbookBranch.Workbook.Repository.GitRepository.Network.Remotes["origin"].PushUrl;
                var credentials = Config.Credentials.Where(c => pushUrl.StartsWith(c.Url)).FirstOrDefault();

                //commit
                workbookBranch.Commit(path, credentials.Username ?? Environment.UserName, credentials.Email);

                //invalidate ribbon
                ribbon.Invalidate();
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

            public void Ribbon_Load(IRibbonUI ribbon)
            {
                this.ribbon = ribbon;
            }

            private void XlApp_WorkbookActivate(Excel.Workbook Wb)
            {
                var path = xlApp.ActiveWorkbook.FullName;
                if (!path.Contains(StagingPath))
                {
                    activeWorkbookBranch = null;
                }
                else
                {
                    activeWorkbookBranch = repositories.GetWorkbookVersionFromPath(GetWorkbookPath(path));
                }
                ribbon.Invalidate();
            }

            public override string GetCustomUI(string RibbonID)
            {
                var str = "<customUI onLoad='Ribbon_Load' xmlns='http://schemas.microsoft.com/office/2006/01/customui'>\n";
                str += "<ribbon>\n";
                str += "<tabs>\n";
                str += "<tab id='xltrail' label='Xltrail'>\n";
                str += "<group id='group1' label='Workbooks'>\n";
                str += "<dynamicMenu id='id-root' label='Workbooks' imageMso='MicrosoftExcel' size='large' getContent='BuildMenu' />\n";
                str += "</group>";
                str += "<group id='group2' label='Save' getVisible='GetWorkbookVisibility'>\n";
                str += "<button id='repositoryName' getLabel='GetRepositoryName' size='normal' imageMso='Info' />\n";
                str += "<button id='workbookName' getLabel='GetWorkbookName' size='normal' imageMso='Info' />\n";
                str += "<button id='branchName' getLabel='GetBranchName' size='normal' imageMso='Info' />\n";
                str += "<button id='discardChangesButton' label='Discard changes' size='normal' imageMso='ReviewRejectChange' getVisible='GetDiscardChangesVisibility' onAction='DiscardChangesButton_Click'/>\n";
                //str += "<button id='commitButton' label='Commit' size='normal' imageMso='FileSave' onAction='CommitButton_Click' />\n";
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
                foreach (var branch in workbook.Branches.Where(b => !b.IsStagingBranch))
                {
                    str += "<button id='" + branch.Id + "' label='" + branch.Description + "' imageMso='MicrosoftExcel' onAction='OpenWorkbook_Click' />\n";
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


            public string GetWorkbookName(IRibbonControl control)
            {
                if (activeWorkbookBranch == null)
                    return "";
                return activeWorkbookBranch.Workbook.Path;
            }

            public string GetRepositoryName(IRibbonControl control)
            {
                if (activeWorkbookBranch == null)
                    return "";
                return activeWorkbookBranch.Workbook.Repository.Name;
            }

            public string GetBranchName(IRibbonControl control)
            {
                if (activeWorkbookBranch == null)
                    return "";
                return activeWorkbookBranch.DisplayName;
            }

            public bool GetWorkbookVisibility(IRibbonControl control)
            {
                if (activeWorkbookBranch == null)
                    return false;
                return true;
            }

            public bool GetDiscardChangesVisibility(IRibbonControl control)
            {
                if (activeWorkbookBranch == null)
                    return false;
                return activeWorkbookBranch.HasConflict;
            }


            public void DiscardChangesButton_Click(IRibbonControl control)
            {
                if (activeWorkbookBranch == null)
                    return;

                var workbookPath = xlApp.ActiveWorkbook.FullName;

                //get `other branch`
                var branch = activeWorkbookBranch.OtherBranch;

                //delete branch and close working copy
                activeWorkbookBranch.Discard();
                xlApp.ActiveWorkbook.Close(false);

                //delete working copy
                File.Delete(workbookPath);

                //re-checkout and re-open
                var stagingBranch = branch.Checkout();
                workbookPath = OpenWorkbook(stagingBranch);
                xlApp.Workbooks.Open(workbookPath);
                ribbon.ActivateTab("xltrail");
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
                    var workbookBranch = repositories.GetWorkbookVersionFromId(control.Id);
                    var stagingBranch = workbookBranch.Checkout();

                    var workbookPath = OpenWorkbook(stagingBranch);
                    if (!stagingBranch.IsStagingBranch)
                    {
                        xlApp.Workbooks.Open(workbookPath, false, true);
                    }
                    else
                    {
                        xlApp.Workbooks.Open(workbookPath);
                    }
                    ribbon.ActivateTab("xltrail");
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
                /*
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
                */
            }


            public void RefreshAll()
            {
                //pull config yaml from remote
                ConfigProvider.PullConfigFromRemote(
                    Config.Repositories,
                    Path.GetDirectoryName(ConfigPathRepositories),
                    Config.Credentials.Where(c => Config.Repositories.StartsWith(c.Url)).FirstOrDefault());

                //load defined repositories from yaml file
                var repositoryConfigs = ConfigProvider.LoadRepositoryConfigsFromFile(ConfigPathRepositories);

                //pull repositories
                foreach(var repositoryConfig in repositoryConfigs)
                {
                    try
                    {
                        var repositoryPath = Path.Combine(RepositoriesPath, repositoryConfig.Alias);
                        var credentials = Config.Credentials.Where(c => repositoryConfig.Url.StartsWith(c.Url)).FirstOrDefault();
                        GitProvider.PullFromRemote(repositoryConfig.Url, repositoryPath, credentials);
                    }
                    catch(Exception ex)
                    {
                        logger.WarnFormat(ex.Message);
                    }
                }

                //load from file system
                repositories = ConfigProvider.LoadRepositoriesFromFilesystem(RepositoriesPath, repositoryConfigs);
            }

        }
    }
}
