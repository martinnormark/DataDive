namespace DataDive
{
	using System;
	using System.Runtime.InteropServices;
	using Microsoft.VisualStudio.Shell;

	/// <summary>
	/// This class implements the tool window exposed by this package and hosts a user control.
	/// </summary>
	/// <remarks>
	/// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
	/// usually implemented by the package implementer.
	/// <para>
	/// This class derives from the ToolWindowPane class provided from the MPF in order to use its
	/// implementation of the IVsUIElementPane interface.
	/// </para>
	/// </remarks>
	[Guid("2df423f4-0422-4279-873e-467b77f77a84")]
	public class FacetsDataWindow : ToolWindowPane
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="FacetsDataWindow"/> class.
		/// </summary>
		public FacetsDataWindow() : base(null)
		{
			this.Caption = "FacetsDataWindow";

			// This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
			// we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
			// the object returned by the Content property.
			this.Content = new FacetsDataWindowControl();
		}

		public void SetData(string json)
			=> ((FacetsDataWindowControl)Content).SetData(json);
	}
}
