using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Tangari.XrmToolBoxExtensions.SolutionImporter
{
    public partial class GTSolutionImporterControlPlugin : PluginControlBase, IHelpPlugin, IPayPalPlugin, IGitHubPlugin
    {
        #region Implementation of the IHelpPlugin interface

        public string HelpUrl
        {
            get { return "https://gennaroeduardotangarisite.wordpress.com/"; }
        }

        #endregion

        #region Implementation of the IPayPalPlugin interface

        public string DonationDescription
        {
            get { return "paypal description"; }
        }

        public string EmailAccount
        {
            get { return "gennarotangari@msn.com"; }
        }

        #endregion

        #region Implementation of the IGitHubPlugin interface

        public string RepositoryName
        {
            get
            {
                return "GTSolutionImporter";
            }
        }

        public string UserName
        {
            get
            {
                return "getangar";
            }
        }

        #endregion  
        public GTSolutionImporterControlPlugin()
        {
            InitializeComponent();
        }
    }
}
