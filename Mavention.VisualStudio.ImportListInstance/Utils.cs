using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE80;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace Mavention.VisualStudio.ImportListInstance {
    public static class Utils {
        /// <summary>
        /// Adds new file to the current Solution/Project and inserts the contents
        /// </summary>
        /// <param name="fileType">File type, eg. General\XML File</param>
        /// <param name="title">File title</param>
        /// <param name="fileContents">File contents</param>
        internal static void CreateNewFile(string fileType, string title, string fileContents) {
            DTE2 dte = Package.GetGlobalService(typeof(SDTE)) as DTE2;
            Document file = dte.ItemOperations.NewFile(fileType, title).Document;
            if (!String.IsNullOrEmpty(fileContents)) {
                TextSelection selection = file.Selection;
                selection.SelectAll();
                selection.Text = "";
                selection.Insert(fileContents);
            }
        }

        internal static Project GetActiveProject() {
            DTE dte = Package.GetGlobalService(typeof(SDTE)) as DTE;
            return GetActiveProject(dte);
        }

        internal static Project GetActiveProject(DTE dte) {
            Project activeProject = null;

            Array activeSolutionProjects = dte.ActiveSolutionProjects as Array;
            if (activeSolutionProjects != null && activeSolutionProjects.Length > 0) {
                activeProject = activeSolutionProjects.GetValue(0) as Project;
            }

            return activeProject;
        }

        internal static void OpenFile(string fileName) {
            DTE dte = Package.GetGlobalService(typeof(SDTE)) as DTE;
            dte.ItemOperations.OpenFile(fileName);
        }
    }
}
