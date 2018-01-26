using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace WordRemoteConverter
{
    public class SingleWordApp : IDisposable
    {
        #region Properties

        private Microsoft.Office.Interop.Word.Application MSdoc = null;
        private int counter = 0;
        private int DISPOSE_AFTER = 500;
        private object missing = Type.Missing;
        private object Unknown = System.Reflection.Missing.Value;
        private object readOnly = false;
        private object FileFormat = WdSaveFormat.wdFormatPDF;
        private object LockComments = false;
        private object AddToRecentFiles = false;
        private object ReadOnlyRecommended = false;
        private object EmbedTrueTypeFonts = true;
        private object SaveNativePictureFormat = false;
        private object SaveFormsData = false;
        private object SaveAsAOCELetter = false;
        private object Encoding = MsoEncoding.msoEncodingUTF8;
        private object InsertLineBreaks = false;
        private object AllowSubstitutions = false;
        private object LineEnding = WdLineEndingType.wdCRLF;
        private object AddBiDiMarks = false;
        private object PrintToFile = true;
        private bool isBusy = false;
        private bool wordavailable = false; 
        private Exception _currentException;

        public bool IsBusy { get
            {
                return isBusy;
            }
        }
        public bool WordIsAvailable
        {
            get
            {
                return wordavailable;
            }
        }

        public Exception GetCurrentException
        {
            get
            {
                return _currentException;
            }
        }

        #endregion


        #region Public Methods

        public SingleWordApp()
        {
            Init(500); 
        }

        public SingleWordApp(int disposeAfter)
        {
            Init(disposeAfter);
        }
         
   

        public string ConvertToPdf(string strFileName, int timeoutInSec, out Exception ex)
        {
            bool timedOut = false;
            isBusy = true;
            if (counter > DISPOSE_AFTER)
            {
                Dispose();
                Init(DISPOSE_AFTER);
            }

            string returnedPdf = "";
            //returnedPdf = _convertToPdf(strFileName, out ex);

            var ts = new CancellationTokenSource();
            CancellationToken ct = ts.Token;
            var task = System.Threading.Tasks.Task.Run(() =>  _convertToPdf(strFileName), ct);
            if (!task.Wait(TimeSpan.FromSeconds(timeoutInSec)))
            {
                try
                {
                    ts.Cancel();
                    timedOut = true;
                }
                catch (Exception ex3)
                {
                }
            }
            returnedPdf = task.Result;
            if (GetCurrentException != null)
                ex = GetCurrentException;
            if (timedOut)
                ex = new Exception("Word timed out to convert");
            ex = null;
            isBusy = false;
            return returnedPdf; 
        }

        public void Dispose()
        {
            try {
                _currentException = null;
                if (MSdoc != null)
                {
                    try { MSdoc.Documents.Close(ref Unknown, ref Unknown, ref Unknown); } catch { }
                    MSdoc.Application.Visible = false;
                    MSdoc.Visible = false;
                    MSdoc.Quit(ref Unknown, ref Unknown, ref Unknown);
                    ReleaseObject(MSdoc);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            } catch (Exception ex){ _currentException = ex; }
            Thread.Sleep(2000);
        }

        #endregion

        #region Private Methods

        private void Init(int disposeAfter)
        {
            try
            {
                _currentException = null;
                if (MSdoc == null)
                {
                    MSdoc = new Microsoft.Office.Interop.Word.Application();
                    ((ApplicationEvents4_Event)MSdoc).Quit += new ApplicationEvents4_QuitEventHandler(_word_application_ApplicationEvents2_Event_Quit);
                    wordavailable = true;
                }
                MSdoc.Application.Visible = false;
                MSdoc.Visible = false;
                MSdoc.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;
                counter = 0;
                DISPOSE_AFTER = disposeAfter;
            }
            catch(Exception ex)
            {
                _currentException = ex; 
                wordavailable = false;
            } 
        }

        private void _word_application_ApplicationEvents2_Event_Quit()
        {
            MSdoc = null;
        }

        private string _convertToPdf(string strFileName)//, out Exception xx)
        {
            try
            {
                _currentException = null;
                object SourceFileName = @strFileName;
                string strFilePath = strFileName.Substring(0, strFileName.LastIndexOf('\\')) + strFileName.Substring(strFileName.LastIndexOf('\\')); //strFilePath.Substring(0, strFilePath.LastIndexOf('.')) + ".txt";
                object newFileName = strFilePath.Substring(0, strFilePath.LastIndexOf('.')) + ".pdf";

                object Source = @SourceFileName;


                MSdoc.Visible = false;
                Microsoft.Office.Interop.Word.Document varDoc = MSdoc.Documents.Open(ref Source, ref Unknown,
                                     ref readOnly, ref Unknown, ref Unknown,
                                     ref Unknown, ref Unknown, ref Unknown,
                                     ref Unknown, ref Unknown, ref Unknown,
                                     ref Unknown, ref Unknown, ref Unknown, ref Unknown, ref Unknown);


                MSdoc.ActiveDocument.SaveAs2(ref newFileName, ref FileFormat, ref missing,
                                           ref missing, ref AddToRecentFiles, ref missing,
                                           ref ReadOnlyRecommended, ref EmbedTrueTypeFonts,
                                           ref SaveNativePictureFormat, ref SaveFormsData,
                                           ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks,
                                           ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks);
                string strReturnFileName = Convert.ToString(newFileName);
                varDoc.Close();

                counter++;
                //xx = null;
                return strReturnFileName;
            }
            catch (Exception ex) {/* xx = ex;*/ _currentException = ex; }
            return null;
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                //obj = null;
            }
            catch (Exception ex)
            {
                //obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion
         
    }
}
