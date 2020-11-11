using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.UI.Grid;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.SqlServer.Management.UI.VSIntegration.Editors;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using Newtonsoft.Json;
using Task = System.Threading.Tasks.Task;

namespace DataDive
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class FacetsDiveCommand
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("bbbf4352-bc56-4939-a805-9a6f56b8d679");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly AsyncPackage package;

		/// <summary>
		/// Initializes a new instance of the <see cref="FacetsDiveCommand"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		/// <param name="commandService">Command service to add command to, not null.</param>
		private FacetsDiveCommand(AsyncPackage package, OleMenuCommandService commandService)
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
		public static FacetsDiveCommand Instance
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
			// Switch to the main thread - the call to AddCommand in FacetsDiveCommand's constructor requires
			// the UI thread.
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
			Instance = new FacetsDiveCommand(package, commandService);
			
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
			string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
			string title = "FacetsDiveCommand";

			//var ss = ServiceProvider.GetServiceAsync(typeof(SVsTextManager)).Result;
			//VSConstants.vsstd

			var dte = Package.GetGlobalService(typeof(SDTE)) as DTE;
			//

			var docs = dte.Documents;
			var win = dte.Windows;

			dte.Application.StatusBar.Text = title;
			var doc = dte.ActiveDocument;
			var type = doc.Type;
			var docwins = doc.Windows;

			var dte2 = (DTE2)ServiceProvider.GetServiceAsync(typeof(DTE)).Result;
			dte2.Events.OutputWindowEvents.PaneUpdated += OutputWindowEvents_PaneUpdated;
			dte2.Events.OutputWindowEvents.PaneAdded += OutputWindowEvents_PaneAdded;

			var w2 = dte2.ActiveDocument.ActiveWindow;
			var dd = w2.DocumentData;

			dte2.Events.CommandEvents.AfterExecute += CommandEvents_AfterExecute;

			var objType = ServiceCache.ScriptFactory.GetType();
			var method1 = objType.GetMethod("GetCurrentlyActiveFrameDocView", BindingFlags.NonPublic | BindingFlags.Instance);
			var Result = method1.Invoke(ServiceCache.ScriptFactory, new object[] { ServiceCache.VSMonitorSelection, false, null });

			var objType2 = Result.GetType();
			var field = objType2.GetField("m_sqlResultsControl", BindingFlags.NonPublic | BindingFlags.Instance);
			var SQLResultsControl = field.GetValue(Result);

			var m_gridResultsPage = GetNonPublicField(SQLResultsControl, "m_gridResultsPage");
			CollectionBase gridContainers = GetNonPublicField(m_gridResultsPage, "m_gridContainers") as CollectionBase;

			foreach (var gridContainer in gridContainers)
			{
				var grid = GetNonPublicField(gridContainer, "m_grid") as GridControl;
				var gridStorage = grid.GridStorage;
				var schemaTable = GetNonPublicField(gridStorage, "m_schemaTable") as DataTable;

				var data = new DataTable();

				for (long i = 0; i < gridStorage.NumRows(); i++)
				{
					var rowItems = new List<object>();

					for (int c = 0; c < schemaTable.Rows.Count; c++)
					{
						var columnName = schemaTable.Rows[c][0].ToString();
						var columnType = schemaTable.Rows[c][12] as Type;

						if (!data.Columns.Contains(columnName))
						{
							data.Columns.Add(columnName, columnType);
						}

						var cellData = gridStorage.GetCellDataAsString(i, c + 1);

						if (cellData == "NULL")
						{
							rowItems.Add(null);

							continue;
						}

						if (columnType == typeof(bool))
						{
							cellData = cellData == "0" ? "False" : "True";
						}

						Console.WriteLine($"Parsing {columnName} with '{cellData}'");
						var typedValue = Convert.ChangeType(cellData, columnType, CultureInfo.InvariantCulture);

						rowItems.Add(typedValue);
					}

					data.Rows.Add(rowItems.ToArray());
				}

				data.AcceptChanges();

				var json = JsonConvert.SerializeObject(data, Formatting.Indented);

				// Get the instance number 0 of this tool window. This window is single instance so this instance
				// is actually the only one.
				// The last flag is set to true so that if the tool window does not exists it will be created.
				ToolWindowPane window = this.package.FindToolWindow(typeof(FacetsDataWindow), 0, true);
				if ((null == window) || (null == window.Frame))
				{
					throw new NotSupportedException("Cannot create tool window");
				}

				var diveWindow = window as FacetsDataWindow;
				diveWindow.SetData(json);

				IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
				Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
			}

			// Show a message box to prove we were here
			//VsShellUtilities.ShowMessageBox(
			//		this.package,
			//		message,
			//		title,
			//		OLEMSGICON.OLEMSGICON_INFO,
			//		OLEMSGBUTTON.OLEMSGBUTTON_OK,
			//		OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

			//Microsoft.SqlServer.Management.UI.Grid.GridControl
			// {Microsoft.SqlServer.Management.UI.VSIntegration.Editors.GridResultsGrid}
		}

		public object GetNonPublicField(object obj, string field)
		{
			FieldInfo f = obj.GetType().GetField(field, BindingFlags.NonPublic | BindingFlags.Instance);

			return f.GetValue(obj);
		}

		private void CommandEvents_AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
		{
		}

		private void OutputWindowEvents_PaneAdded(OutputWindowPane pPane)
		{
		}

		private void OutputWindowEvents_PaneUpdated(OutputWindowPane pPane)
		{
			
		}
	}
}
