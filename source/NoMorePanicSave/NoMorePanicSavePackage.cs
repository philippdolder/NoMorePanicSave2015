//------------------------------------------------------------------------------
// <copyright file="NoMorePanicSavePackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace PhilippDolder.NoMorePanicSave
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Runtime.InteropServices;
    using EnvDTE;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell;
    using NLog;
    using NLog.Config;
    using NLog.Targets;
    using Process = System.Diagnostics.Process;

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
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(NoMorePanicSavePackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string)]
    public sealed class NoMorePanicSavePackage : Package
    {
        private IntPtr otherApplicationFocusedHookHandle;
        private IntPtr currentInstanceFocusedHookHandle;

        private WindowsEventHooker.WinEventDelegate otherApplicationFocusedHandlerReference;
        private WindowsEventHooker.WinEventDelegate currentInstanceFocusedHandlerReference;

        private readonly Logger logger;

        private bool enabled;
        private bool visualStudioClosing;
        private SolutionEvents solutionEvents;

        /// <summary>
        /// NoMorePanicSavePackage GUID string.
        /// </summary>
        public const string PackageGuidString = "20615f30-3b3c-4b74-a552-cec8cd21650c";

        /// <summary>
        /// Initializes a new instance of the <see cref="NoMorePanicSavePackage"/> class.
        /// </summary>
        public NoMorePanicSavePackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
#if DEBUG
            this.logger = LogManager.GetLogger("NoMorePanicSave");
            var configuration = new LoggingConfiguration();

            var target = new DebuggerTarget();
            configuration.AddTarget("console", target);

            var rule = new LoggingRule("*", LogLevel.Trace, target);
            configuration.LoggingRules.Add(rule);

            this.logger.Factory.Configuration = configuration;
#else
            this.logger = LogManager.CreateNullLogger();
#endif
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            this.logger.Debug("Initializing...");

            var visualStudioProcess = Process.GetCurrentProcess();
            this.otherApplicationFocusedHandlerReference = this.HandleOtherApplicationFocused;
            this.currentInstanceFocusedHandlerReference = this.HandleCurrentInstanceFocused;
            this.otherApplicationFocusedHookHandle = WindowsEventHooker.SetWinEventHook(3, 3, IntPtr.Zero, this.otherApplicationFocusedHandlerReference, 0, 0, SetWinEventHookFlags.WINEVENT_OUTOFCONTEXT | SetWinEventHookFlags.WINEVENT_SKIPOWNPROCESS);
            this.currentInstanceFocusedHookHandle = WindowsEventHooker.SetWinEventHook(3, 3, IntPtr.Zero, this.currentInstanceFocusedHandlerReference, (uint)visualStudioProcess.Id, 0, SetWinEventHookFlags.WINEVENT_OUTOFCONTEXT);

            try
            {
                var dte = (DTE)this.GetService(typeof(DTE));

                this.solutionEvents = dte.Events.SolutionEvents;
                this.solutionEvents.BeforeClosing += this.HandleBeforeClosingSolution;
            }
            catch (Exception e)
            {
                this.logger.Error(e, "Could not hook visual studio closing event.");
            }

            this.logger.Debug("Initialized...");
        }

        protected override void Dispose(bool disposing)
        {
            this.logger.Debug("disposing package");

            WindowsEventHooker.UnhookWinEvent(this.currentInstanceFocusedHookHandle);
            this.currentInstanceFocusedHandlerReference = null;

            WindowsEventHooker.UnhookWinEvent(this.otherApplicationFocusedHookHandle);
            this.otherApplicationFocusedHandlerReference = null;

            base.Dispose(disposing);

            this.logger.Debug("disposed package");
        }

        private void HandleBeforeClosingSolution()
        {
            this.visualStudioClosing = true;
            this.logger.Debug("solution before closing. `visualStudioClosing = true`");
        }

        private void HandleCurrentInstanceFocused(IntPtr hWinEventHook, uint eventType,
                IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            this.logger.Debug("current instance was focused. `enabled = true`");
            this.enabled = true;
        }

        private void HandleOtherApplicationFocused(
                IntPtr hWinEventHook, uint eventType,
                IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            this.logger.Debug("other application was focused. EventHook: `{0}`; EventType: `{1}`; WindowHandle: `{2}`; ObjectId: `{3}`; ChildId: `{4}`; Thread: `{5}`; Time: `{6}`;", hWinEventHook, eventType, hWnd, idObject, idChild, dwEventThread, dwmsEventTime);

            if (this.enabled && !this.visualStudioClosing)
            {
                this.SaveAll();

                this.enabled = false;
            }
        }

        private void SaveAll()
        {
            try
            {
                DTE dte = (DTE)this.GetService(typeof(DTE));

                dte.ExecuteCommand("File.SaveAll");

                this.logger.Debug("saved solution.");

            }
            catch (Exception e)
            {
                this.logger.Error(e, "Saving all files failed.");
            }
        }
    }
}
