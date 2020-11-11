using System;
using System.Runtime.InteropServices;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace DataDive
{
	/// <summary>
	/// This is the class that implements the package exposed by this assembly.
	/// </summary>
	/// <remarks>
	/// <para>
	/// The minimum requirement for a class to be considered a valid package for Visual Studio
	/// is to implement the IVsPackage interface and register itself with the shell.
	/// This package uses the helper classes defined inside the Managed Package Framework (MPF)
	/// to do it: it derives from the Package class that provides the implementation of the
	/// IVsPackage interface and uses the registration attributes defined in the framework to
	/// register itself and its components with the shell. These attributes tell the pkgdef creation
	/// utility what data to put into .pkgdef file.
	/// </para>
	/// <para>
	/// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
	/// </para>
	/// </remarks>
	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[Guid(DataDivePackage.PackageGuidString)]
	[ProvideMenuResource("Menus.ctmenu", 1)]
	[ProvideToolWindow(typeof(FacetsDataWindow))]
	public sealed class DataDivePackage : AsyncPackage
	{
		/// <summary>
		/// DataDivePackage GUID string.
		/// </summary>
		public const string PackageGuidString = "30cb9448-cb4a-4d75-adf3-f82075e0453e";

		public CommandEvents QueryExecuteEvent { get; private set; }

		#region Package Members

		/// <summary>
		/// Initialization of the package; this method is called right after the package is sited, so this is the place
		/// where you can put all the initialization code that rely on services provided by VisualStudio.
		/// </summary>
		/// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
		/// <param name="progress">A provider for progress updates.</param>
		/// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
		protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
		{
			// When initialized asynchronously, the current thread may be a background thread at this point.
			// Do any initialization that requires the UI thread after switching to the UI thread.
			await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
		    await FacetsDiveCommand.InitializeAsync(this);



			var dte2 = (DTE2)GetGlobalService(typeof(DTE));
			var command = dte2.Commands.Item("Query.Execute");

			QueryExecuteEvent = dte2.Events.get_CommandEvents(command.Guid, command.ID);
			QueryExecuteEvent.BeforeExecute += QueryExecuteEvent_BeforeExecute;
			QueryExecuteEvent.AfterExecute += QueryExecuteEvent_AfterExecute;
		  await FacetsDataWindowCommand.InitializeAsync(this);
		}

		private void QueryExecuteEvent_AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			var dte2 = (DTE2)GetGlobalService(typeof(DTE));
			var dte = (DTE)GetGlobalService(typeof(DTE));
			var doc = dte2.ActiveDocument;

			var ok = dte2.ActiveDocument.ActiveWindow.ObjectKind;
			//var ss = dte2.ActiveDocument.ActiveWindow.Object("m_sqlResultsControl");

			var text = doc.Object("TextDocument") as TextDocument;
			var t = doc.Object();
			var g = doc.Object("m_sqlResultsControl");
		}

		private void QueryExecuteEvent_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
		{
			var dte2 = (DTE2)GetGlobalService(typeof(DTE));
		}

		#endregion
	}
}
