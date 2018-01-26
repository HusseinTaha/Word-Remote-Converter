using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordRemoteConverter
{
    public class RemoteConverter : MarshalByRefObject
    {
        private SingleWordApp wordApp;
        private bool wordavailable = false;
        private bool checkedword = false;

        public RemoteConverter()
        { 
        }

        public bool WordIsAvailable()
        {
            if (!checkedword)
            {
                try
                {
                    if(wordApp == null)
                    {
                        using (SingleWordApp tempWordApp = new SingleWordApp())
                        {
                            wordavailable = tempWordApp.WordIsAvailable;
                        }
                    }
                    else
                    {
                        wordavailable = wordApp.WordIsAvailable;
                    }
                }
                catch
                {
                    wordavailable = false;
                }
                checkedword = true;
            }
            return wordavailable;
        }

        public string Convert(string source, out Exception ex)
        {
            if (wordApp == null)
                wordApp = new SingleWordApp(30);
            if (wordApp != null && !wordApp.IsBusy)
                return wordApp.ConvertToPdf(source, 15, out ex);
            ex = new Exception("WordApp is null or Word is busy");
            return null;
        }

        public void Dispose()
        {
            if (wordApp != null && !wordApp.IsBusy)
            {
                wordApp.Dispose();
                wordApp = null;
            }
        }
    }
}
