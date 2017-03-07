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
using Vertex.Omiga.Supervisor.Service.Types.Entities;
using Vertex.Omiga.Supervisor.Service.Types.Requests;
using Vertex.Omiga.Supervisor.Service.Types.Responses;
using Vertex.Omiga.Supervisor.Service.Types.Faults;

namespace Vertex.Omiga.Supervisor.Client
{
    public partial class FixedGlobalParameter : Form
    {
        private GlobalParameter globalParameter;
        private bool isEdit = false;

        public FixedGlobalParameter()
        {
            InitializeComponent();
            globalParameter = new GlobalParameter();
        }

        public FixedGlobalParameter(GlobalParameter globalParameterInput)
        {
            InitializeComponent();
            globalParameter = globalParameterInput;
            isEdit = true;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            SaveData();
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            globalParameter = null;
            this.Close();
        }

        private void FixedGlobalParameter_Shown(object sender, EventArgs e)
        {
            PopulateData();
        }

        private void PopulateData()
        {
            if (!string.IsNullOrEmpty(globalParameter.Name))
            {
                NameTextBox.Text = globalParameter.Name;
                StartDate.Value = globalParameter.StartDate;
                Description.Text = globalParameter.Description;
                AmountTextBox.Text = globalParameter.ValueAmount.ToString();
                StringTextBox.Text = globalParameter.ValueString;
                if (globalParameter.ValueBoolean.HasValue)
                {
                    BooleanCombo.SelectedIndex = (globalParameter.ValueBoolean ?? false) ? 1 : 2;
                }
                else
                {
                    BooleanCombo.SelectedIndex = 0;
                }
                Percentage.Text = globalParameter.ValuePercentage.ToString();
                MaximumAmount.Text = globalParameter.ValueMaximumAmount.ToString();
                NameTextBox.ReadOnly = true;
                StartDate.Enabled = false;
            }
            else
            {
                StartDate.Value = DateTime.Now;
                BooleanCombo.SelectedIndex = 0;
            }
        }

        private void SaveData()
        {
            globalParameter.Name = NameTextBox.Text;
            globalParameter.StartDate = StartDate.Value;
            globalParameter.Description = Description.Text;
            globalParameter.ValueAmount = AssignNullableDouble(AmountTextBox.Text);
            if (StringTextBox.Text.Length > 0)
            {
                globalParameter.ValueString = StringTextBox.Text;
            }
            else
            {
                globalParameter.ValueString = null;
            }
            switch (BooleanCombo.SelectedIndex)
            {
                case 0: //<Select>
                    globalParameter.ValueBoolean = null;
                    break;
                case 1: //True
                    globalParameter.ValueBoolean = true;
                    break;
                case 2: //False
                    globalParameter.ValueBoolean = false;
                    break;
            }
            globalParameter.ValuePercentage = AssignNullableDouble(Percentage.Text);
            globalParameter.ValueMaximumAmount = AssignNullableDouble(MaximumAmount.Text);
            
            GlobalParameter returnedGlobal;

            if (isEdit)
            {
                returnedGlobal = UpdateGlobalParameter();
            }
            else
            {
                returnedGlobal = CreateGlobalParameter();
            }

            if (returnedGlobal != null)
            {
                globalParameter = returnedGlobal;
            }
        }

        private GlobalParameter CreateGlobalParameter()
        {
            ChannelFactory<IOmigaSupervisorService> factory = new ChannelFactory<IOmigaSupervisorService>("NetTcpBinding_IOmigaSupervisorService");
            IOmigaSupervisorService serviceProxy = factory.CreateChannel();

            try
            {
                CreateGlobalParameterRequest request = new CreateGlobalParameterRequest();
                ClientUtility.PopulateRequestContext(request.Context);
                request.EntityToCreate = globalParameter;

                CreateGlobalParameterResponse response = serviceProxy.CreateGlobalParameter(request);
                return response.NewEntity;
            }

            catch (FaultException<OmigaSupervisorFault> fault)
            {
                MessageBox.Show(string.Format("{0}: {1} ({2})", fault.Reason, fault.Detail.FaultMessage, fault.Detail.FaultGuid));
                return null;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private GlobalParameter UpdateGlobalParameter()
        {
            ChannelFactory<IOmigaSupervisorService> factory = new ChannelFactory<IOmigaSupervisorService>("NetTcpBinding_IOmigaSupervisorService");
            IOmigaSupervisorService serviceProxy = factory.CreateChannel();

            try
            {
                UpdateGlobalParameterRequest request = new UpdateGlobalParameterRequest();
                ClientUtility.PopulateRequestContext(request.Context);
                request.EntityToUpdate = globalParameter;

                UpdateGlobalParameterResponse response = serviceProxy.UpdateGlobalParameter(request);
                return response.UpdatedEntity;
            }

            catch (FaultException<OmigaSupervisorFault> fault)
            {
                MessageBox.Show(string.Format("{0}: {1} ({2})", fault.Reason, fault.Detail.FaultMessage, fault.Detail.FaultGuid));
                return null;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public GlobalParameter ReturnedGlobalParameter
        {
            get { return globalParameter; }
            set { globalParameter = value; }
        }

        private double? AssignNullableDouble(string textAmount)
        {
            double? value = null;
            double parsedValue;

            if (double.TryParse(textAmount, out parsedValue))
            {
                value = parsedValue;
            }
            return value;
        }
    }
}