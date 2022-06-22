using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;

namespace SetMultipleStartupProjects
{
    [Command(PackageIds.MyCommand)]
    internal sealed class MyCommand
    {
        private DTE2 dte;

        private MyCommand(DTE2 dte, IMenuCommandService commandService)
        {
            this.dte = dte;

            if (commandService != null)
            {
                var menuCommandID = new CommandID(PackageGuids.SetMultipleStartupProjects, PackageIds.MyCommand);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        public static async Task InitializeAsync(IAsyncServiceProvider package)
        {
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as IMenuCommandService;

            var dte = await package.GetServiceAsync(typeof(DTE)) as DTE2;

            new MyCommand(dte, commandService);
        }

        public async Task<List<object>> GetStartUpProjectsAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var startupProjects = (object[])this.dte.Solution.SolutionBuild.StartupProjects;

            return (startupProjects ?? Enumerable.Empty<object>()).ToList();
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            Task.Run(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                var projects = (await VS.Solutions.GetActiveItemsAsync());

                var startupProjects = await GetStartUpProjectsAsync();

                this.dte.Solution.SolutionBuild.StartupProjects = projects.Select(x => x.FullPath).ToArray<object>();
            });
        }
    }
}