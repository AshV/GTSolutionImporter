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
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using System.IO;

namespace Tangari.XrmToolBoxExtensions.SolutionImporter
{
    public partial class GTSolutionImporterControlPlugin : PluginControlBase, IHelpPlugin, IPayPalPlugin, IGitHubPlugin
    {
        #region Private parameters

        private Boolean isSolutionSelected = false;

        #endregion

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

        #region Constructors
        public GTSolutionImporterControlPlugin()
        {
            InitializeComponent();
        }

        #endregion

        #region Toolstrip buttons

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool(); // PluginBaseControl method that notifies the XrmToolBox that the user wants to close the plugin
            // Override the ClosingPlugin method to allow for any plugin specific closing logic to be performed (saving configs, canceling close, etc...)
        }

        private void tsbAbout_Click(object ender, EventArgs e)
        {

            MessageBox.Show("GT Solution Importer - Version " + typeof(GTSolutionImporter).Assembly.GetName().Version + "\n\n(c)Coypright by Gennaro Eduardo Tangari", "GT Solution Importer");
        }

        private void tsbCancel_Click(object sender, EventArgs e)
        {
            CancelWorker(); // PluginBaseControl method that calls the Background Workers CancelAsync method.

            MessageBox.Show("Cancelled");
        }

        private void tsbClear_Click(object sender, EventArgs e)
        {
            
        }

        #endregion

        #region Interface

        private void btnGetOrgs_Click(object sender, EventArgs e)
        {
            Type thisType = this.GetType();
            Type serviceType = this.Service.GetType();

            String OrganizationServiceUrl = (String)(((thisType.GetProperty("ConnectionDetail").GetValue(this, null)).GetType()).GetProperty("OrganizationServiceUrl")).GetValue(thisType.GetProperty("ConnectionDetail").GetValue(this, null), null);
            String DiscoveryServiceUrl = OrganizationServiceUrl.Substring(0, OrganizationServiceUrl.LastIndexOf('/')) + "/Discovery.svc";
            Boolean IsCrmOnline = (Boolean)((((thisType.GetProperty("ConnectionDetail").GetValue(this, null)).GetType()).GetProperty("UseOnline")).GetValue(thisType.GetProperty("ConnectionDetail").GetValue(this, null), null));

            if (IsCrmOnline)
            {
                MessageBox.Show("GT Solution Importer currently support only On-Premise deployment only!", "GT Solution Importer");
                lstOrgs.DataSource = null;

                return;
            }

            ClientCredentials credentials = (ClientCredentials)serviceType.GetProperty("ClientCredentials").GetValue(this.Service, null);
            
            using (var discoveryProxy = new DiscoveryServiceProxy(new Uri(DiscoveryServiceUrl), null, credentials, null))
            {
                discoveryProxy.Authenticate();

                // Get all Organizations using Discovery Service

                RetrieveOrganizationsRequest retrieveOrganizationsRequest =
                new RetrieveOrganizationsRequest()
                {
                    AccessType = EndpointAccessType.Default,
                    Release = OrganizationRelease.Current
                };

                RetrieveOrganizationsResponse retrieveOrganizationsResponse =
                (RetrieveOrganizationsResponse)discoveryProxy.Execute(retrieveOrganizationsRequest);
               
                if (retrieveOrganizationsResponse.Details.Count > 0)
                {
                    Dictionary<string, string> collections = new Dictionary<string, string>();

                    foreach (OrganizationDetail orgInfo in retrieveOrganizationsResponse.Details)
                    {
                        collections.Add(orgInfo.OrganizationId.ToString(), orgInfo.FriendlyName);
                    }


                    lstOrgs.DataSource = new BindingSource(collections, null);
                    lstOrgs.DisplayMember = "Value";
                    lstOrgs.ValueMember = "Key";
                    lstOrgs.SelectedIndex = 0;
                }
            }
        }

        private void btnSelectSolution_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Title = "GT Solution Importer";
            openFile.Filter = "ZIP|*.zip";
            openFile.Multiselect = false;
            openFile.ShowDialog();

            if (openFile.FileName != "")
            {
                txtSolutionPath.Text = openFile.FileName;
                lblSolution.Text = "Solution to import: " + openFile.FileName.Substring(openFile.FileName.LastIndexOf(@"\") + 1);

                isSolutionSelected = true;
            }
            else
            {
                isSolutionSelected = false;
                lblSolution.Text = "Solution not selected";

            }
        }


        #endregion

        private void btnImportSolution_Click(object sender, EventArgs e)
        {
            if (lstOrgs.Items.Count > 0 && lstOrgs.SelectedItems.Count > 0)
            {
                if (isSolutionSelected)
                {
                    if (System.IO.File.Exists(txtSolutionPath.Text))
                    {
                        if (System.IO.Path.GetExtension(txtSolutionPath.Text).ToUpper() == ".ZIP")
                        {
                            for (int i = 0; i < lstOrgs.SelectedItems.Count; i++)
                            {                                
                                System.Collections.Generic.KeyValuePair<string, string> obj =
                                    (System.Collections.Generic.KeyValuePair<string, string>)lstOrgs.SelectedItems[i];

                                MessageBox.Show(obj.Key + " - " + obj.Value, "GT Solution Importer");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Before to proceed please select a valid solution file!", "GT Solution Importer");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Before to proceed please select a valid solution file!", "GT Solution Importer");
                }
            }
            else
            {
                MessageBox.Show("Before to proceed please select at least one target organization!", "GT Solution Importer");
            }
        }
    }
}