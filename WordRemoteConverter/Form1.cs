using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Windows.Forms;

namespace WordRemoteConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //initialize remoting
            RemotingConfiguration.Configure("WordRemoteConverter.exe.config", false);
            RemotingConfiguration.RegisterWellKnownServiceType(new RemoteConverter().GetType(), "RemoteConverter", WellKnownObjectMode.Singleton);


            //check if word 2013 is available
            try
            {
                using (SingleWordApp wordapp = new SingleWordApp())
                {
                    if (wordapp.WordIsAvailable)
                    {
                        lbl_wordinstalled.ForeColor = Color.Green;
                        lbl_wordinstalled.Text = "available";
                    }
                    else
                    {
                        lbl_wordinstalled.ForeColor = Color.Red;
                        lbl_wordinstalled.Text = "not available";
                    }
                }

            }
            catch
            {
                lbl_wordinstalled.ForeColor = Color.Red;
                lbl_wordinstalled.Text = "not available";
            }
        }
    }
}
