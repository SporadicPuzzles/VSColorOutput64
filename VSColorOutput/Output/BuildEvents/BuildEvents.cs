﻿using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading;
using VSColorOutput.FindResults;
using VSColorOutput.Output.TimeStamp;
using VSColorOutput.State;

// ReSharper disable PrivateFieldCanBeConvertedToLocalVariable

#pragma warning disable 649

namespace VSColorOutput.Output.BuildEvents
{
    public class BuildEvents : IVsUpdateSolutionEvents2
    {
        private readonly Stopwatch _buildDurationStopwatch = new Stopwatch();

        private DTE2 _dte2;
        private Events _events;
        private DTEEvents _dteEvents;
        private int _initialized;
        private List<string> _projectsBuildReport;
        private IVsSolution _solutionService;
        private IVsSolutionBuildManager2 _sbm;
        private string _buildConfig = string.Empty;

        public bool StopOnBuildErrorEnabled { get; set; }
        public bool ShowElapsedBuildTimeEnabled { get; set; }
        public string ElapsedTimeFormatString { get; set; }
        public bool ShowBuildReport { get; set; }
        public bool ShowDebugWindowOnDebug { get; set; }
        public bool ShowTimeStamps { get; set; }
        public DateTime DebugStartTime { get; private set; }
        public bool ShowDonation { get; set; }
        public static string SolutionPath { get; private set; }
        public bool ShowWarningCount { get; set; }

        public void Initialize(IServiceProvider serviceProvider)
        {
            if (Interlocked.CompareExchange(ref _initialized, 1, 0) == 1) return;

#pragma warning disable VSSDK006 // Check services exist
            _solutionService = serviceProvider.GetService(typeof(SVsSolution)) as IVsSolution;
            _sbm = serviceProvider.GetService(typeof(SVsSolutionBuildManager)) as IVsSolutionBuildManager2;

            // NOTE this class is never disposed, so we don't track the cookie here and don't unadvise on disposal
            ErrorHandler.ThrowOnFailure(_sbm!.AdviseUpdateSolutionEvents(this, out _));

            _dte2 = serviceProvider.GetService(typeof(DTE)) as DTE2;
#pragma warning restore VSSDK006 // Check services exist
            if (_dte2 != null)
            {
                // These event sources have to be rooted or the GC will collect them.
                // https://social.msdn.microsoft.com/Forums/en-US/vsx/thread/fd2f9108-1df3-4d96-a65d-67a69347ca27
                _events = _dte2.Events;
                _dteEvents = _events.DTEEvents;

                _dteEvents.ModeChanged += OnModeChanged;
            }

            _projectsBuildReport = new List<string>();

            // If not started from Test session
            if (_solutionService != null)
            {
                if (IsSolutionLoaded())
                {
                    SolutionOpened();
                }

                Microsoft.VisualStudio.Shell.Events.SolutionEvents.OnAfterOpenSolution += (s, e) => SolutionOpened();
                Microsoft.VisualStudio.Shell.Events.SolutionEvents.OnAfterCloseSolution += (s, e) => SolutionClosed();
            }

            Settings.SettingsUpdated += (sender, args) => LoadSettings();
            LoadSettings();

            WarningsCount.Initialize();
        }

        private bool IsSolutionLoaded()
        {
            _solutionService.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out object value);

            return value is bool isSolOpen && isSolOpen;
        }

        public void SolutionOpened()
        {
            SolutionPath = Path.GetDirectoryName(_dte2.Solution.FullName);
            LoadSettings();
        }

        public void SolutionClosed()
        {
            SolutionPath = null;
        }

        private void LoadSettings()
        {
            var settings = Settings.Load();
            StopOnBuildErrorEnabled = settings.EnableStopOnBuildError;
            ShowElapsedBuildTimeEnabled = settings.ShowElapsedBuildTime;
            ElapsedTimeFormatString = settings.TimeStampElapsed ?? @"hh\:mm\:ss\.fff";
            ShowBuildReport = settings.ShowBuildReport;
            ShowDebugWindowOnDebug = settings.ShowDebugWindowOnDebug;
            ShowTimeStamps = settings.ShowTimeStamps;
            ShowDonation = !settings.SuppressDonation;
            ShowWarningCount = settings.ShowWarningCount;
        }

        /// <summary>Build is starting.</summary>
        public int UpdateSolution_Begin(ref int pfCancelUpdate)
        {
            _projectsBuildReport.Clear();
            _buildDurationStopwatch.Restart();
            return VSConstants.S_OK;
        }

        /// <summary>Build is complete.</summary>
        public int UpdateSolution_Done(int fSucceeded, int fModified, int fCancelCommand)
        {
            bool BuildSucceeded = fSucceeded != 0;

            OutputWindowPane buildOutputPane = null;
            // ReSharper disable once LoopCanBeConvertedToQuery
            foreach (OutputWindowPane pane in _dte2.ToolWindows.OutputWindow.OutputWindowPanes)
            {
                if (pane.Guid == VSConstants.OutputWindowPaneGuid.BuildOutputPane_string)
                {
                    buildOutputPane = pane;
                    break;
                }
            }

            if (buildOutputPane == null)
            {
                return VSConstants.S_OK;
            }

            if (ShowBuildReport)
            {
                buildOutputPane.OutputString($"{Environment.NewLine}Projects build report:{Environment.NewLine}");
                buildOutputPane.OutputString($"  Status    | Project [{_buildConfig}]{Environment.NewLine}");
                buildOutputPane.OutputString($" -----------|---------------------------------------------------------------------------------------------------{Environment.NewLine}");
                foreach (var reportItem in _projectsBuildReport)
                {
                    buildOutputPane.OutputString(reportItem + Environment.NewLine);
                }
            }

            if (ShowWarningCount && BuildSucceeded)
            {
                WarningsCount.DisplayWarnings(buildOutputPane);
            }

            // Ensure the stop watch has started. A case where the start event hasn't fired has been observed.
            if (ShowElapsedBuildTimeEnabled && _buildDurationStopwatch.IsRunning)
            {
                // The next build must start it again, so that we don't end up timing across multiple builds.
                _buildDurationStopwatch.Stop();

                var elapsed = _buildDurationStopwatch.Elapsed;
                var time = elapsed.ToString(ElapsedTimeFormatString);
                var buildTime = DateTime.Now.ToString(CultureInfo.CurrentCulture);
                var timeElapsed = $"Build time {time}";
                var endedAt = $"Build ended at {buildTime}";
                buildOutputPane.OutputString($"{Environment.NewLine}{timeElapsed}{Environment.NewLine}");
                buildOutputPane.OutputString($"{endedAt}{Environment.NewLine}");
            }

            if (ShowDonation)
            {
                buildOutputPane.OutputString($"{Environment.NewLine}++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
                buildOutputPane.OutputString($"{Environment.NewLine}+++                 Please consider donating to VSColorOutput                    +++");
                buildOutputPane.OutputString($"{Environment.NewLine}+++                       https://mike-ward.net/donate/                          +++");
                buildOutputPane.OutputString($"{Environment.NewLine}+++            (this message can be turned off in the settings panel)            +++");
                buildOutputPane.OutputString($"{Environment.NewLine}++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++{Environment.NewLine}");
            }

            return VSConstants.S_OK;
        }

        /// <summary>A project's build is done.</summary>
        public int UpdateProjectCfg_Done(IVsHierarchy pHierProj, IVsCfg pCfgProj, IVsCfg pCfgSln, uint dwAction, int fSuccess, int fCancel)
        {
            var success = fSuccess != 0;

            if (StopOnBuildErrorEnabled && !success)
            {
                const string cancelBuildCommand = "Build.Cancel";
                _dte2.ExecuteCommand(cancelBuildCommand);
            }

            if (ShowWarningCount)
                WarningsCount.RetrieveBuildLog(pHierProj, pCfgProj);

            if (ShowBuildReport)
            {
                ErrorHandler.ThrowOnFailure(pHierProj.GetProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_Name, out var project));
                pCfgProj.get_DisplayName(out var displayName);
                _buildConfig = displayName;

                _projectsBuildReport.Add($"  {(success ? "Succeeded" : "Failed   ")} | {project}");
            }

            return VSConstants.S_OK;
        }

        private void OnModeChanged(vsIDEMode lastMode)
        {
            if (lastMode == vsIDEMode.vsIDEModeDesign)
            {
                DebugStartTime = DateTime.Now;

                if (ShowDebugWindowOnDebug || ShowTimeStamps && !TimeStampMarginProvider.Initialized)
                {
                    ActivateDebugOutputWindow();
                }
            }
        }

        private void ActivateDebugOutputWindow()
        {
            _dte2.ToolWindows.OutputWindow.Parent.Activate();
            foreach (OutputWindowPane pane in _dte2.ToolWindows.OutputWindow.OutputWindowPanes)
            {
                if (pane.Guid == VSConstants.OutputWindowPaneGuid.DebugPane_string)
                {
                    pane.Activate();
                    break;
                }
            }
        }

        public int UpdateSolution_StartUpdate(ref int pfCancelUpdate) => VSConstants.S_OK;

        public int UpdateSolution_Cancel() => VSConstants.S_OK;

        public int OnActiveProjectCfgChange(IVsHierarchy pIVsHierarchy) => VSConstants.S_OK;

        public int UpdateProjectCfg_Begin(IVsHierarchy pHierProj, IVsCfg pCfgProj, IVsCfg pCfgSln, uint dwAction, ref int pfCancel) => VSConstants.S_OK;
    }
}