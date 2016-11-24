using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class UserControl1 : UserControl
    {
        private readonly Uri initialUri;
        public UserControl1()
        {
            InitializeComponent();
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);

            this.webBrowser2.Navigating += WebBrowser2_Navigating;
            initialUri = new Uri(uriCodeBase, @"HTMLPage1.html");
            this.webBrowser2.Url = initialUri;

        }

        /// <summary>
        /// Checks if the URL is navigating to a page I am expecting namely the close event handler otherwise just silently ignore.  It will only
        /// allow nativation to the initial uri.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WebBrowser2_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            e.Cancel = initialUri != e.Url;
            if (e.Url.Scheme == "myapp") {
                if (e.Url.PathAndQuery == "/closeme")
                {
                    Globals.ThisAddIn.toggleCustomTaskPane();
                }
            }
        }
    }
}
