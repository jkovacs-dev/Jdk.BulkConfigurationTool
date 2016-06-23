using Jdk.BulkConfigurationTool.AppCode;
using System;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

namespace Jdk.BulkConfigurationTool
{
    public partial class MainControl : PluginControlBase
    {
        public MainControl()
        {
            InitializeComponent();
        }

        public static void AddItem(ListBox box, string item)
        {
            MethodInvoker miAddItem = delegate
            {
                box.Items.Add(item);
            };

            if (box.InvokeRequired)
            {
                box.Invoke(miAddItem);
            }
            else
            {
                miAddItem();
            }
        }


        private void AddLogItem(object sender, string message)
        {
            AddItem(listLog, message);
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using(var ofd = new OpenFileDialog())
            {
                ofd.Title = "Specify the file to process";
                ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                ofd.FilterIndex = 1;
                ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                ofd.RestoreDirectory = true;

                if (ofd.ShowDialog(ParentForm) == DialogResult.OK)
                {
                    txtFilePath.Text = ofd.FileName;
                }
            }
        }

        private void txtFilePath_TextChanged(object sender, EventArgs e)
        {
            tsbImport.Enabled = tsbDelete.Enabled = (txtFilePath.Text.Length != 0 && Service != null);
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }

        private void tsbImport_Click(object sender, EventArgs e)
        {
            if (txtFilePath.Text.Length == 0)
                return;

            listLog.Items.Clear();
            var originalTimeout = ConnectionDetail.ServiceClient.OrganizationServiceProxy.Timeout;
            ConnectionDetail.ServiceClient.OrganizationServiceProxy.Timeout = new TimeSpan(0, 30, 00);

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Importing Configuration...",
                AsyncArgument = new object[] { txtFilePath.Text },
                Work = (bw, evt) =>
                {
                    var filePath = ((object[])evt.Argument)[0].ToString();
                    var ie = new ConfigurationFileImporter(filePath);
                    ie.RaiseSuccess += AddLogItem;
                    ie.RaiseError += AddLogItem;
                    var data = ie.Import();
                    var processor = new CreateCrmDataProcessor(Service, data);
                    processor.RaiseSuccess += AddLogItem;
                    processor.RaiseError += AddLogItem;
                    processor.ProcessData();
                },
                PostWorkCallBack = evt => {
                    ConnectionDetail.ServiceClient.OrganizationServiceProxy.Timeout = originalTimeout;
                }
            });

        }

        private void tsbDelete_Click(object sender, EventArgs e)
        {
            if (txtFilePath.Text.Length == 0)
                return;

            listLog.Items.Clear();
            var originalTimeout = ConnectionDetail.ServiceClient.OrganizationServiceProxy.Timeout;
            ConnectionDetail.ServiceClient.OrganizationServiceProxy.Timeout = new TimeSpan(0, 30, 00);
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Removing Configuration...",
                AsyncArgument = new object[] { txtFilePath.Text },
                Work = (bw, evt) =>
                {
                    var filePath = ((object[])evt.Argument)[0].ToString();
                    var ie = new ConfigurationFileImporter(filePath);
                    ie.RaiseSuccess += AddLogItem;
                    ie.RaiseError += AddLogItem;
                    var data = ie.Import();
                    var processor = new RemoveCrmDataProcessor(Service, data);
                    processor.RaiseSuccess += AddLogItem;
                    processor.RaiseError += AddLogItem;
                    processor.ProcessData();
                },
                PostWorkCallBack = evt => {
                    ConnectionDetail.ServiceClient.OrganizationServiceProxy.Timeout = originalTimeout;
                }
            });
        }

        private void tsbExport_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Title = "Select where to save the file";
                sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                sfd.FilterIndex = 1;
                sfd.FileName = "BulkConfiguration.xlsx";
                sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                sfd.RestoreDirectory = true;
                sfd.OverwritePrompt = true;

                if (sfd.ShowDialog(ParentForm) != DialogResult.OK)
                {
                    return;
                }

                listLog.Items.Clear();
                if (txtFilePath.Text.Length == 0)
                {
                    txtFilePath.Text = sfd.FileName;
                }
                WorkAsync(new WorkAsyncInfo
                {
                    Message = "Exporting Configuration File Template...",
                    AsyncArgument = new object[] { sfd.FileName },
                    Work = (bw, evt) =>
                    {
                        var filePath = ((object[])evt.Argument)[0].ToString();
                        var ee = new ConfigurationFileExporter(filePath);
                        ee.RaiseSuccess += AddLogItem;
                        ee.RaiseError += AddLogItem;
                        ee.Export();
                    },
                    PostWorkCallBack = evt => { }
                });
            }
        }
    }
}
