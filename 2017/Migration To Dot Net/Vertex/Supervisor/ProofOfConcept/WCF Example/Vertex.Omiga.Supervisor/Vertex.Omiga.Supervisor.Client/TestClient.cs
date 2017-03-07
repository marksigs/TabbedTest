using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.ServiceModel;
using System.ServiceModel.Channels;

using Vertex.Omiga.Supervisor.Service.Interface;
using Vertex.Omiga.Supervisor.Service.Types.Requests;
using Vertex.Omiga.Supervisor.Service.Types.Responses;
using Vertex.Omiga.Supervisor.Service.Types.Faults;
using Vertex.Omiga.Supervisor.Service.Types.Entities;

namespace Vertex.Omiga.Supervisor.Client
{
    public partial class TestClient : Form
    {
        public TestClient()
        {
            InitializeComponent();
        }

        #region Control events
        private void menuFileExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            switch (e.Node.Name)
            {
                case "rootNode":
                    statusMessage.Text = string.Empty;
                    summaryListView.Visible = false;
                    break;
                case "globalParameters":
                    statusMessage.Text = CallService();
                    break;
            }
        }


        private void summaryListView_DoubleClick(object sender, EventArgs e)
        {
            EditListItem();
        }

        private void EditListItem()
        {
            using (FixedGlobalParameter form = new FixedGlobalParameter((GlobalParameter) summaryListView.SelectedItems[0].Tag))
            {
                form.ShowDialog();
                GlobalParameter updatedGlobal = form.ReturnedGlobalParameter;
                if (updatedGlobal != null)
                {
                    ListViewItem item = summaryListView.SelectedItems[0];
                    item.Text = updatedGlobal.Name;
                    item.Tag = updatedGlobal;
                    item.SubItems[1].Text = updatedGlobal.StartDate.ToShortDateString();
                    item.SubItems[2].Text = updatedGlobal.Description;
                    item.SubItems[3].Text = updatedGlobal.ValueAmount.ToString();
                    item.SubItems[4].Text = updatedGlobal.ValueString;
                    item.SubItems[5].Text = updatedGlobal.ValueBoolean.ToString();
                    item.SubItems[6].Text = updatedGlobal.ValuePercentage.ToString();
                    item.SubItems[7].Text = updatedGlobal.ValueMaximumAmount.ToString();
                }
            }
        }

        private void menuEditAddNewItem_Click(object sender, EventArgs e)
        {
            AddNewGlobalParameter();
        }

        private void deleteContextMenuItem_Click(object sender, EventArgs e)
        {
            DeleteListItem();
        }

        private void editContextMemuItem_Click(object sender, EventArgs e)
        {
            EditListItem();
        }

        private void newContextMenuItem_Click(object sender, EventArgs e)
        {
            AddNewGlobalParameter();
        }

        private void summaryListView_KeyUp(object sender, KeyEventArgs e)
        {
            if (summaryListView.SelectedItems.Count > 0)
            {
                switch (e.KeyCode)
                {
                    case Keys.Enter:
                        EditListItem();
                        break;
                    case Keys.Back:
                    case Keys.Delete:
                        DeleteListItem();
                        break;
                }
            }
        }

        #endregion

        #region Helpers
        /// <summary>
        /// Call service and return response description
        /// </summary>
        /// <returns>Text suitable for display</returns>
        private string CallService()
        {
            ChannelFactory<IOmigaSupervisorService> factory = new ChannelFactory<IOmigaSupervisorService>("NetTcpBinding_IOmigaSupervisorService");
            IOmigaSupervisorService serviceProxy = factory.CreateChannel();

            try
            {
                GetGlobalParametersRequest request = new GetGlobalParametersRequest();
                ClientUtility.PopulateRequestContext(request.Context);

                GetGlobalParametersResponse response = serviceProxy.GetGlobalParameters(request);

                DisplayGlobalParameters(response);

                TimeSpan elapsed = response.Context.SentAt.Subtract(response.Context.RequestContext.ReceivedAt);

                return string.Format("Service returned {0} in {1}ms", response.GlobalParameters.Length, elapsed.TotalMilliseconds);
            }

            catch (FaultException<OmigaSupervisorFault> fault)
            {
                return string.Format("{0}: {1} ({2})", fault.Reason, fault.Detail.FaultMessage, fault.Detail.FaultGuid);
            }

            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private void DisplayGlobalParameters(GetGlobalParametersResponse response)
        {
            this.Cursor = Cursors.WaitCursor;

            //Put the list items into an array so they can be added to the listview in one go
            int i = 0;
            ListViewItem[] items = new ListViewItem[response.GlobalParameters.Length];
            foreach (GlobalParameter globalParameter in response.GlobalParameters)
            {
                ListViewItem item = new ListViewItem(globalParameter.Name, 0);
                item.Tag = globalParameter;
                item.SubItems.Add(globalParameter.StartDate.ToShortDateString());
                item.SubItems.Add(globalParameter.Description);
                item.SubItems.Add(globalParameter.ValueAmount.ToString());
                item.SubItems.Add(globalParameter.ValueString);
                item.SubItems.Add(globalParameter.ValueBoolean.ToString());
                item.SubItems.Add(globalParameter.ValuePercentage.ToString());
                item.SubItems.Add(globalParameter.ValueMaximumAmount.ToString());

                items[i++] = item;
            }

            summaryListView.BeginUpdate();
            
            summaryListView.Items.Clear();
            summaryListView.Columns.Clear();
            summaryListView.Columns.Add("Name", 180, HorizontalAlignment.Left);
            summaryListView.Columns.Add("Start Date", 70, HorizontalAlignment.Left);
            summaryListView.Columns.Add("Description", 200, HorizontalAlignment.Left);
            summaryListView.Columns.Add("Amount", 70, HorizontalAlignment.Right);
            summaryListView.Columns.Add("String", 150, HorizontalAlignment.Left);
            summaryListView.Columns.Add("Boolean", 70, HorizontalAlignment.Left);
            summaryListView.Columns.Add("Percentage", 70, HorizontalAlignment.Right);
            summaryListView.Columns.Add("Maximum Amount", 70, HorizontalAlignment.Right);
            
            summaryListView.Items.AddRange(items);
            summaryListView.Visible = true;
            summaryListView.EndUpdate();
            this.Cursor = Cursors.Default;
        }

        #endregion

        private void DeleteListItem()
        {
            if (summaryListView.SelectedItems.Count > 0)
            {
                DialogResult result = MessageBox.Show("Are you sure you want to delete the item " + summaryListView.SelectedItems[0].Text + "?", 
                                                "Deletion Confirmation",
                                                MessageBoxButtons.YesNo, 
                                                MessageBoxIcon.Exclamation);
                if (result == DialogResult.Yes)
                {
                    GlobalParameter globalToDelete = (GlobalParameter)summaryListView.SelectedItems[0].Tag;
                    GlobalParameterIdentity identity = new GlobalParameterIdentity();
                    identity.Name = globalToDelete.Name;
                    identity.StartDate = globalToDelete.StartDate;
                    identity.TimeStamp = globalToDelete.TimeStamp;
                    if (DeleteGlobalParameter(identity))
                    {
                        summaryListView.SelectedItems[0].Remove();
                    }
                }
            }
        }

        private bool DeleteGlobalParameter(GlobalParameterIdentity globalParameterToDelete)
        {
            if (globalParameterToDelete != null)
            {
                ChannelFactory<IOmigaSupervisorService> factory = new ChannelFactory<IOmigaSupervisorService>("NetTcpBinding_IOmigaSupervisorService");
                IOmigaSupervisorService serviceProxy = factory.CreateChannel();

                try
                {
                    DeleteGlobalParameterRequest request = new DeleteGlobalParameterRequest();
                    ClientUtility.PopulateRequestContext(request.Context);
                    request.EntityToDelete = globalParameterToDelete;

                    DeleteGlobalParameterResponse response = serviceProxy.DeleteGlobalParameter(request);
                    return true;
                }

                catch (FaultException<OmigaSupervisorFault> fault)
                {
                    MessageBox.Show(string.Format("{0}: {1} ({2})", fault.Reason, fault.Detail.FaultMessage, fault.Detail.FaultGuid));
                    return false;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void AddNewGlobalParameter()
        {
            using (FixedGlobalParameter form = new FixedGlobalParameter())
            {
                form.ShowDialog();
                GlobalParameter newGlobal = form.ReturnedGlobalParameter;
                if (newGlobal != null)
                {
                    ListViewItem item = new ListViewItem(newGlobal.Name, 0);
                    item.Tag = newGlobal;
                    item.SubItems.Add(newGlobal.StartDate.ToShortDateString());
                    item.SubItems.Add(newGlobal.Description);
                    item.SubItems.Add(newGlobal.ValueAmount.ToString());
                    item.SubItems.Add(newGlobal.ValueString);
                    item.SubItems.Add(newGlobal.ValueBoolean.ToString());
                    item.SubItems.Add(newGlobal.ValuePercentage.ToString());
                    item.SubItems.Add(newGlobal.ValueMaximumAmount.ToString());
                    summaryListView.Items.Add(item);
                }
            }
        }
    }

}
