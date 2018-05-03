using Microsoft.TeamFoundation.VersionControl.Client;
using System;
using System.IO;
using System.Linq;

namespace Microsoft.Operations.TeamFoundationServer
{
    /// <summary>
    /// Borrowed from
    /// http://www.vlaquest.com/2012/07/tfs-api-a-tiny-class-to-help-with-tfs-workspace-creation-and-cleanup/
    /// Also refer MSDN: http://msdn.microsoft.com/en-us/magazine/jj883959.aspx
    /// </summary>
    /// <summary>
    /// Tiny helper class for using temporary workspaces. It will create and remove a workspace for
    /// you, so that you don't have to manage it in code.
    /// </summary>
    /// <example>
    /// TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(new Uri(serverName));
    /// tfs.EnsureAuthenticated(); VersionControlServer vcs = (VersionControlServer)tfs.GetService(typeof(VersionControlServer));
    ///
    /// using (var tempWorkspace = new TfsTempWorkspace(vcs, "ServiceAccountWorkspace",
    /// string.Format(@"{0}\{1}", Environment.UserDomainName, Environment.UserName))) { string
    /// localPath = tempWorkspace.Map(string.Format("$/{0}/Reconcillation_Files", projectName),
    /// FileSystem.WorkingFolder("Gallacake", "Temp"));
    ///
    /// if (vcs.ServerItemExists(expectedLocationInVersionControl, ItemType.File)) { // If the same
    /// named item already exists, this will do a write-over. // Technically this will never happen
    /// by the logic of the previous processing, // but in situations where the item is being
    /// replayed, this is useful. tempWorkspace.Workspace.PendEdit(finalSpreadsheet.FullName); } else
    /// { tempWorkspace.Workspace.PendAdd(finalSpreadsheet.FullName); }
    ///
    /// int changesetNumber =
    /// tempWorkspace.Workspace.CheckIn(tempWorkspace.Workspace.GetPendingChanges(), versionControlCheckinComment);
    ///
    /// sw.WriteLine("{0} Version Control: Changeset # {1} File: {2}", DateTime.Now, changesetNumber,
    /// localPath); }
    /// </example>
    public class TfsTempWorkspace : IDisposable
    {
        private VersionControlServer _vcs;
        private Workspace _workspace;

        /// <summary>
        /// Creates a workspace. If a workspace with the same name exists, it will be deleted and
        /// recreated! For best results, give your workspace a relatively unique name. ALSO Note:
        /// when the item gets disposed, all local files are also removed (i.e. because they are all committed)
        /// </summary>
        /// <param name="deleteExistingFilesPriorToMapping">
        /// Indicates whether you want to flush the files in the nominated directory. Set to true if
        /// you need it to start clean.
        /// </param>
        public TfsTempWorkspace(VersionControlServer vcs, string workspaceName, string userName, bool deleteExistingFilesPriorToMapping = false)
        {
            _vcs = vcs;
            _workspace = _vcs.QueryWorkspaces(workspaceName, null, Environment.MachineName).FirstOrDefault(); // using 'null' will force it to query all users. This is actually the only way it will find the existing workspace.

            if (_workspace != null)
            {
                _workspace.Delete();
            }
            _workspace = _vcs.CreateWorkspace(workspaceName);

            if (deleteExistingFilesPriorToMapping)
            {
                CleanDirectory();
            }
        }

        /// <summary>
        /// Give access to many properties of the workspace
        /// </summary>
        public Workspace Workspace
        {
            get
            {
                return _workspace;
            }
        }

        public static void DeleteDirectory(string target_dir)
        {
            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }

            Directory.Delete(target_dir, false);
        }

        public void CleanDirectory()
        {
            string rootPath = Path.Combine(Path.GetTempPath(), _workspace.Name);
            if (Directory.Exists(rootPath))
            {
                DeleteDirectory(rootPath);
            }
        }

        /// <summary>
        /// All local files are deleted, and the workspace is then removed from the server
        /// </summary>
        public void Dispose()
        {
            CleanDirectory();
            _workspace.Delete();
        }

        /// <summary>
        /// Adds a mapping to the workspace. The local folders and files will be created under the
        /// user TEMP directory.
        /// </summary>
        /// <param name="serverPath">Full path on server, starting from root, ex: $/MyTP/MyFolder</param>
        /// <param name="localRelativePath">A relative path inside the local workspace folder structure</param>
        /// <returns>The local full path</returns>
        public string Map(string serverPath, string localRelativePath)
        {
            string localPath = Path.Combine(Path.GetTempPath(), _workspace.Name, localRelativePath);
            _workspace.Map(serverPath, localPath);
            return localPath;
        }
    }
}