using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace GitStalebranchCleaner
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CleanStaleBranchesCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ae43e998-b926-407c-8595-2607a254ebc3");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CleanStaleBranchesCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private CleanStaleBranchesCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CleanStaleBranchesCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in CleanStaleBranchesCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CleanStaleBranchesCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                // Get the solution directory
                var dte = Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                if (dte?.Solution == null || string.IsNullOrEmpty(dte.Solution.FullName))
                {
                    ShowMessage("No solution is currently open.", OLEMSGICON.OLEMSGICON_WARNING);
                    return;
                }

                string solutionDir = System.IO.Path.GetDirectoryName(dte.Solution.FullName);

                // Execute the Git commands
                CleanStaleBranches(solutionDir);
            }
            catch (Exception ex)
            {
                ShowMessage($"Error cleaning stale branches: {ex.Message}", OLEMSGICON.OLEMSGICON_CRITICAL);
            }
        }

        private void CleanStaleBranches(string repoPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var output = new System.Text.StringBuilder();

            // First, prune remote references
            output.AppendLine("Pruning remote references...");
            ExecuteGitCommand(repoPath, "fetch --prune", output);

            // Get list of stale branches
            string branchOutput = ExecuteGitCommand(repoPath, "branch -vv");
            var staleBranches = new System.Collections.Generic.List<string>();

            using (var reader = new System.IO.StringReader(branchOutput))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Contains(": gone]"))
                    {
                        // Extract branch name (first token after trimming)
                        string branchName = line.Trim().Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries)[0];
                        if (branchName.StartsWith("*"))
                            branchName = line.Trim().Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries)[1];

                        staleBranches.Add(branchName);
                    }
                }
            }

            if (staleBranches.Count == 0)
            {
                output.AppendLine("\nNo stale branches found.");
                ShowMessage(output.ToString(), OLEMSGICON.OLEMSGICON_INFO);
                return;
            }

            output.AppendLine($"\nFound {staleBranches.Count} stale branch(es) :");

            // Delete each stale branch
            foreach (var branch in staleBranches)
            {
                output.AppendLine($"  Deleting: {branch}");
                ExecuteGitCommand(repoPath, $"branch -D {branch}", output);
            }

            output.AppendLine("\nCleanup complete!");
            ShowMessage(output.ToString(), OLEMSGICON.OLEMSGICON_INFO);
        }

        private string ExecuteGitCommand(string workingDirectory, string arguments, System.Text.StringBuilder output = null)
        {
            var processInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "git",
                Arguments = arguments,
                WorkingDirectory = workingDirectory,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = System.Diagnostics.Process.Start(processInfo))
            {
                string result = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(error) && output != null)
                {
                    output.AppendLine($"Error: {error}");
                }

                return result;
            }
        }

        private void ShowMessage(string message, OLEMSGICON icon)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                "Clean Stale Branches",
                icon,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
