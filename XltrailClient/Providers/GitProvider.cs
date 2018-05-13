using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xltrail.Client.Models;
using Xltrail.Client.Models.Config;

namespace Xltrail.Client.Providers
{
    public class GitProvider
    {
        private static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static void EnsureStagingBranches(string path)
        {
            using (var repo = new LibGit2Sharp.Repository(path))
            {
                foreach (var branch in repo.Branches.Where(b => b.IsRemote))
                {
                    //we need to duplicate the branches
                    var localBranchName = branch.FriendlyName.Replace(branch.RemoteName + "/", "");
                    var stagingBranchName = localBranchName + Branch.StagingBranchPrefix;
                    if (repo.Branches[stagingBranchName] == null)
                    {
                        repo.Branches.Update(
                            repo.Branches.Add(stagingBranchName, branch.Tip),
                            b => b.TrackedBranch = branch.CanonicalName);
                    }
                }
            }
        }

        public static void PullFromRemote(string url, string path, Credentials credentials)
        {
            logger.InfoFormat("Start repository sync: {0}", path);
            if (!Directory.Exists(path))
            {
                logger.Info("Repository does not exist");
                logger.InfoFormat("Clone from {0}", url);
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
                logger.Info("Repository exists");
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
                        logger.InfoFormat("Fetch repository from {0}: {1}", remote.Name, remote.Url);
                        LibGit2Sharp.Commands.Fetch(repository, remote.Name, refSpecs, fetchOptions, "");
                    }

                    var branchesUpdated = 0;
                    var branchesCreated = 0;
                    foreach (var branch in repository.Branches.Where(b => b.IsRemote))
                    {
                        var localBranchName = branch.FriendlyName.Replace(branch.RemoteName + "/", "");
                        if (repository.Branches[localBranchName] == null)
                        {
                            logger.InfoFormat("Branch does not exist: {0}", localBranchName);
                            logger.InfoFormat("Create branch from {0}: {1}", branch.FriendlyName, branch.Tip.Sha.Substring(0, 7));
                            repository.Branches.Update(
                                repository.Branches.Add(localBranchName, branch.Tip),
                                b => b.TrackedBranch = branch.CanonicalName);
                            branchesCreated++;
                        }

                        var localBranch = repository.Branches[localBranchName];
                        if (localBranch.Tip.Sha != branch.Tip.Sha)
                        {
                            //check out branch and reset
                            logger.InfoFormat("Branch {0} is behind: {1}", localBranchName, localBranch.Tip.Sha.Substring(0, 7));
                            logger.InfoFormat("Fast-forward to {0}", branch.Tip.Sha.Substring(0, 7));
                            LibGit2Sharp.Commands.Checkout(repository, localBranchName);
                            repository.Reset(LibGit2Sharp.ResetMode.Hard, branch.Tip);
                            branchesUpdated++;
                        }
                        else
                        {
                            logger.InfoFormat("Branch {0} is in sync: {1}", localBranch.FriendlyName, localBranch.Tip.Sha.Substring(0, 7));
                        }
                    }
                    logger.InfoFormat("Created branches: {0}", branchesCreated);
                    logger.InfoFormat("Updated branches: {0}", branchesUpdated);
                    logger.InfoFormat("Repository sync ready: {0}", path);
                }
            }
        }
    }
}
