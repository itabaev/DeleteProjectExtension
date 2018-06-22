using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using DeleteProjectExtension.Properties;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace DeleteProjectExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class DeleteProjectCommand
    {
        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Verify the current thread is the UI thread - the call to AddCommand in DeleteProjectCommand's constructor requires
            // the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new DeleteProjectCommand(package, commandService);
        }

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("dcab2cd5-10ac-48ff-b398-34b8e2dc4f5b");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteProjectCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private DeleteProjectCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static DeleteProjectCommand Instance
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
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            DTE dte = (DTE)await this.ServiceProvider.GetServiceAsync(typeof(DTE));
            dynamic[] activeProjects = (dynamic[])dte.ActiveSolutionProjects;
            string message = activeProjects.Length == 1
                ? string.Format(Resources.DeletingProjectMessage, $"'{activeProjects.First().Name}'")
                : string.Format(Resources.DeletingProjectsMessage, string.Join(", ", activeProjects.Select(x => $"'{x.Name}'")));

            // Show a message box to prove we were here
            var messageBoxResult = VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                null,
                OLEMSGICON.OLEMSGICON_WARNING,
                OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            if (messageBoxResult == 1)
            {
                Dictionary<dynamic, Exception> exceptions = new Dictionary<dynamic, Exception>();
                bool error = false;
                string solutionPath = Path.GetDirectoryName(dte.Solution.FileName)?.TrimEnd('\\');
                foreach (var project in activeProjects)
                {
                    try
                    {
                        string projectPath = Path.GetDirectoryName(project.FileName)?.TrimEnd('\\');

                        dte.Solution.Remove(project);

                        if (string.Equals(solutionPath, projectPath, StringComparison.OrdinalIgnoreCase))
                        {
                            exceptions.Add(project, new Exception(Resources.SolutionAndProjectSameDirectoryMessage));
                        }
                        else
                        {
                            Directory.Delete(projectPath, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(project, ex);
                        error = true;
                    }
                }

                if (exceptions.Any())
                {
                    string resultMessage = string.Join(Environment.NewLine, exceptions.Select(x => $"'{x.Key.Name}': {x.Value.Message}"));
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        resultMessage,
                        null,
                        error ? OLEMSGICON.OLEMSGICON_CRITICAL : OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                }
            }
        }
    }
}
