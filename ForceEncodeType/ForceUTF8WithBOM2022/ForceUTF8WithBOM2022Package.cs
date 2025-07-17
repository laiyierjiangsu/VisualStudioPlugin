using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Ude;
using Task = System.Threading.Tasks.Task;

namespace ForceUTF8WithBOM2022
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
    [Guid(ForceUTF8WithBOM2022Package.PackageGuidString)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class ForceUTF8WithBOM2022Package : AsyncPackage
    {
        /// <summary>
        /// ForceUTF8WithBOM2022Package GUID string.
        /// </summary>
        public const string PackageGuidString = "406f07d5-2c08-4fe9-a64b-b7080635141f";

        private static readonly Guid OutputPaneGuid = new Guid("78E31F73-D9CE-4E36-8937-D6BB18E7AA68");

        private IVsOutputWindowPane outputPane;

        private DocumentEvents documentEvents;

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
            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            Assumes.Present(dte);

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var outputWindow = GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            Guid outputPaneGuid = OutputPaneGuid;
            outputWindow.GetPane(ref outputPaneGuid, out outputPane);
            if (outputPane == null)
            {
                outputWindow.CreatePane(ref outputPaneGuid, "ForceUTF8WithBOM", 1, 1);
                outputWindow.GetPane(ref outputPaneGuid, out outputPane);
            }

            documentEvents = dte.Events.DocumentEvents;
            documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;

            Output("Initialized");
        }

        private void Output(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            outputPane.OutputStringThreadSafe(message + Environment.NewLine);
        }

        void DocumentEvents_DocumentSaved(Document document)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Output("Document saved " + document.Kind + " " + document.FullName);

            if (document.Kind != "{8E7B96A8-E33D-11D0-A6D5-00C04FB67F6A}")
            {
                Output("Not text file, skipped");
                return;
            }

            var path = document.FullName;

            var withBOM = false;
            var encoding = DetectFileEncoding(path, out withBOM);
            Output("Detected " + encoding + " " + encoding.CodePage + " withBOM " + withBOM);
            if (encoding == Encoding.UTF8 && withBOM)
            {
                Output("Already UTF-8 with BOM, skipped");
                return;
            }

            using (var stream = File.OpenRead(path))
            {
                using (var reader = new StreamReader(stream, encoding))
                {
                    string text = reader.ReadToEnd();
                    reader.Close();
                    stream.Close();
                    byte[] utf8Bytes = Encoding.Convert(encoding, Encoding.UTF8, encoding.GetBytes(text));
                    File.WriteAllText(path, Encoding.UTF8.GetString(utf8Bytes), new UTF8Encoding(true, true));
                }
            }

            Output("Converted to UTF-8 with BOM");
        }

        private Encoding DetectFileEncoding(string filePath, out bool withBOM)
        {
            using (var stream = File.OpenRead(filePath))
            {
                var detector = new CharsetDetector();
                detector.Feed(stream);
                detector.DataEnd();
                ThreadHelper.ThrowIfNotOnUIThread();
                Output("Detected " + detector.Charset + " " + detector.Confidence);

                var charset = "gb18030";
                if (detector.Charset != null && (detector.Confidence >= 1 || detector.Charset.ToLower().StartsWith("utf")))
                {
                    charset = detector.Charset;
                }

                var bom = new byte[3];
                stream.Position = 0;
                stream.Read(bom, 0, 3);
                withBOM = bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF;

                return Encoding.GetEncoding(charset);
            }
        }

        #endregion
    }
}
