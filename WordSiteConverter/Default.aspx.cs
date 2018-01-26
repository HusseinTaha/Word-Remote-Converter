using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WordSiteConverter
{
    public partial class Default : System.Web.UI.Page
    {
        private WordRemoteConverter.RemoteConverter converter;
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                //try to get the remoting-object
                converter = (WordRemoteConverter.RemoteConverter)Activator.GetObject(typeof(WordRemoteConverter.RemoteConverter),
                    "http://localhost:8989/RemoteConverter");

                if (!converter.WordIsAvailable())
                {
                    lbl_error.Text = "Word 2013 not available on server!";
                    FileUpload1.Visible = false;
                    btnConvert.Visible = false;
                }

            }
            catch
            {
                //Remoteserver not available
                lbl_error.Text = "PDFConverter is not running!";
                FileUpload1.Visible = false;
                btnConvert.Visible = false;
            }
        }

        protected void btnConvert_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                if (FileUpload1.FileName.Contains(".doc"))
                {
                    Exception ex = new Exception();
                    string sourcefile = Path.GetTempPath() + FileUpload1.FileName;
                    FileUpload1.SaveAs(sourcefile);
                    string destination = converter.Convert(sourcefile, out ex);
                    lbl_error.Text = destination;
                }
            }
        }

        protected void btnDispose_Click(object sender, EventArgs e)
        {
            converter.Dispose();
        }
    }
}