using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSColorOutput.FindResults
{
    internal class WarningsCount
    {
        private static HashSet<string> _BuildLogFiles = new HashSet<string>();
        private static Regex _WarningsRegex = new Regex(@"(\W|^)(warning|warn)\W", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static void RetrieveBuildLog(IVsHierarchy pHierProj, IVsCfg pCfgProj)
        {
            if (pHierProj is IVsBuildPropertyStorage propertyStorage)
            {
                string[] configNames = new string[1];
                if (ErrorHandler.Succeeded(pCfgProj.get_DisplayName(out configNames[0])))
                {
                    string intermediateOutputPath;
                    int hr = propertyStorage.GetPropertyValue("IntermediateOutputPath", configNames[0], (uint)_PersistStorageType.PST_PROJECT_FILE, out intermediateOutputPath);

                    if (ErrorHandler.Succeeded(hr) && !string.IsNullOrEmpty(intermediateOutputPath))
                    {
                        string projectDir;
                        if (ErrorHandler.Succeeded(pHierProj.GetCanonicalName((uint)VSConstants.VSITEMID_ROOT, out projectDir)))
                        {
                            ErrorHandler.ThrowOnFailure(pHierProj.GetProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_Name, out var project));
                            string project_log_file = $"{project}.log";

                            string absoluteIntermediateOutputPath = Path.Combine(Path.GetDirectoryName(projectDir), intermediateOutputPath);
                            string projectLogFilePath = Path.Combine(absoluteIntermediateOutputPath, project_log_file);

                            if (File.Exists(projectLogFilePath))
                                _BuildLogFiles.Add(projectLogFilePath);
                            else
                                throw new Exception($"Build logs {project_log_file} not found at {projectLogFilePath}.");
                        }
                    }
                }
            }
        }

        public static void DisplayWarnings(OutputWindowPane buildOutputPane)
        {
            List<Tuple<uint, string>> WarningsByProject = new List<Tuple<uint, string>>();

            int LongestProjectName = 0;
            uint TotalWarningsCount = 0;
            foreach (string ProjectLogPath in _BuildLogFiles)
            {
                uint ProjectWarningsCount = (uint)_WarningsRegex.Matches(File.ReadAllText(ProjectLogPath)).Count;

                if (ProjectWarningsCount > 0)
                {
                    string ProjectName = Path.GetFileNameWithoutExtension(ProjectLogPath);

                    if (ProjectName.Length > LongestProjectName)
                        LongestProjectName = ProjectName.Length;

                    WarningsByProject.Add(Tuple.Create(ProjectWarningsCount, ProjectName));

                    TotalWarningsCount += ProjectWarningsCount;
                }
            }

            if (TotalWarningsCount > 0)
            {
                string ProjectsString = "Projects";
                if (ProjectsString.Length > LongestProjectName)
                    LongestProjectName = ProjectsString.Length;

                string WarningsDisplay = $"{Environment.NewLine}";
                WarningsDisplay += $"  {ProjectsString.PadRight(LongestProjectName)} | Warnings{Environment.NewLine}";
                WarningsDisplay += $"--{"---".PadRight(LongestProjectName, '-')} | -----------{Environment.NewLine}";

                foreach (var WarningsForProject in WarningsByProject)
                    WarningsDisplay += $"  {WarningsForProject.Item2.PadRight(LongestProjectName)} | {WarningsForProject.Item1}{Environment.NewLine}";

                WarningsDisplay += $"--{"---".PadRight(LongestProjectName, '-')} | -----------{Environment.NewLine}";
                WarningsDisplay += $"  {"Total".PadRight(LongestProjectName)} | {TotalWarningsCount}";

                buildOutputPane.OutputString(WarningsDisplay);
            }
        }
    }
}
