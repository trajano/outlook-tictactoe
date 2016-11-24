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
        private readonly string initialUri;
        public UserControl1()
        {
            InitializeComponent();
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);

            this.webControl1.WebView.BeforeNavigate += WebView_BeforeNavigate;
            initialUri = new Uri(uriCodeBase, @"HTMLPage1.html").ToString();
            this.webControl1.WebView.Url = initialUri;

        }

        private void WebView_BeforeNavigate(object sender, EO.WebBrowser.BeforeNavigateEventArgs e)
        {
            e.Cancel = initialUri != e.NewUrl;
            Uri theUri = new Uri(e.NewUrl);
            if (theUri.Scheme == "myapp")
            {
                if (theUri.PathAndQuery == "/closeme")
                {
                    Globals.ThisAddIn.toggleCustomTaskPane();
                }
            }

        }
    }
}
