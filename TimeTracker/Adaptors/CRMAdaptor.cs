using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using TimeTracker.Common;
using TimeTracker.Entity;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;


namespace TimeTracker.Adaptors
{
    public class CRMAdaptor
    {
        public Uri OrganizationUri;

        ClientCredentials DeviceCrednetials;
        ClientCredentials Credentials;
        OrganizationServiceProxy _serviceProxy;

        IOrganizationService _service;

        #region Public Methods
        public CRMAdaptor()
        {
            GetCRMConnection();
        }

        public bool? UpdateTaskInCRM(string ticketnumber, TimeItem timeItem)
        {
            try
            {
                Microsoft.Xrm.Sdk.Entity followup = _service.Retrieve("task", (Guid)timeItem.crmTaskId, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));

                Microsoft.Crm.Sdk.Messages.SetStateRequest req = new Microsoft.Crm.Sdk.Messages.SetStateRequest();
                req.EntityMoniker = new Microsoft.Xrm.Sdk.EntityReference("task", followup.Id);
                req.State = new Microsoft.Xrm.Sdk.OptionSetValue(0);
                req.Status = new Microsoft.Xrm.Sdk.OptionSetValue((-1));
                _service.Execute(req);

                if (timeItem.title != null)
                {
                    followup["subject"] = timeItem.title.Replace(ticketnumber, "");
                }

                if (timeItem.description != null)
                {
                    followup["description"] = timeItem.description;
                }

                if (timeItem.isBillable == true && timeItem.isBillable != null)
                {
                    followup["actualdurationminutes"] = (int)TimeSpan.Parse(timeItem.time).TotalMinutes;
                    followup["hsal_nonbillableduration"] = 0;
                }
                else
                {
                    followup["actualdurationminutes"] = 0;
                    followup["hsal_nonbillableduration"] = (int)TimeSpan.Parse(timeItem.time).TotalMinutes;
                }

                followup["actualstart"] = DateTime.UtcNow;

                _service.Update(followup);

                req.EntityMoniker = new Microsoft.Xrm.Sdk.EntityReference("task", followup.Id);
                req.State = new Microsoft.Xrm.Sdk.OptionSetValue(1);
                req.Status = new Microsoft.Xrm.Sdk.OptionSetValue((5));
                _service.Execute(req);

                return true;
            }
            catch(Exception ex)
            {
                //log error
                return false;
            }
        }

        public Guid? CreateTaskToCRMIncident(string ticketnumber, TimeItem timeItem)
        {
            try
            {
                Microsoft.Xrm.Sdk.Query.QueryExpression GetCasesByTicketNumber = new Microsoft.Xrm.Sdk.Query.QueryExpression { EntityName = "incident", ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet(true) };
                GetCasesByTicketNumber.Criteria.AddCondition("ticketnumber", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, ticketnumber);
                Microsoft.Xrm.Sdk.EntityCollection CaseResults = _service.RetrieveMultiple(GetCasesByTicketNumber);

                if (CaseResults.Entities.Count < 1)
                {
                    return null;
                }

                Microsoft.Xrm.Sdk.Entity followup = new Microsoft.Xrm.Sdk.Entity("task");

                if (timeItem.title != null)
                {
                    followup["subject"] = timeItem.title.Replace(ticketnumber, "");
                }

                if (timeItem.description != null)
                {
                    followup["description"] = timeItem.description;
                }

                if (timeItem.isBillable == true && timeItem.isBillable != null)
                {
                    followup["actualdurationminutes"] = (int)TimeSpan.Parse(timeItem.time).TotalMinutes;
                    followup["hsal_nonbillableduration"] = 0;
                }
                else
                {
                    followup["actualdurationminutes"] = 0;
                    followup["hsal_nonbillableduration"] = (int)TimeSpan.Parse(timeItem.time).TotalMinutes;
                }

                followup["actualstart"] = DateTime.UtcNow;

                followup["regardingobjectid"] = CaseResults.Entities[0].ToEntityReference();

                Guid taskId = _service.Create(followup);

                Microsoft.Crm.Sdk.Messages.SetStateRequest req = new Microsoft.Crm.Sdk.Messages.SetStateRequest();
                req.EntityMoniker = new Microsoft.Xrm.Sdk.EntityReference("task", taskId);
                req.State = new Microsoft.Xrm.Sdk.OptionSetValue(1);
                req.Status = new Microsoft.Xrm.Sdk.OptionSetValue((5));
                _service.Execute(req);

                return taskId;
            }
            catch (Exception ex)
            {
                //log error
                return null;
            }
        }

        public List<TimeItem> CreateCaseTask(List<TimeItem> timeItems)
        {

            Regex r = new Regex(Constants.HLSCaseNumberRegex, RegexOptions.IgnoreCase); 

            if (this._service == null)
            {
                return timeItems;
            }

            foreach (TimeItem item in timeItems)
            {
                String[] itemTitleParsed = item.title.Split((string[])null, StringSplitOptions.RemoveEmptyEntries);
                foreach (string word in itemTitleParsed)
                {
                    if (r.Match(word).Success && item.isCRMSubmitted == false)
                    {
                        if (item.crmTaskId != null)
                        {
                            item.isCRMSubmitted = UpdateTaskInCRM(word, item);
                        }
                        else
                        {
                            item.crmTaskId = CreateTaskToCRMIncident(word, item);
                            item.isCRMSubmitted = true;
                        }                   
                    }
                }
            }
            return timeItems;
        }

        #endregion Public Methods

        #region Private Methods
        private void GetCRMConnection()
        {
            DeviceCrednetials = null;
            Credentials = new ClientCredentials();
            try
            {
                XmlTextReader reader = new XmlTextReader("Connection.xml");

                string element = "";
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        element = reader.Name;
                    }
                    else if (reader.NodeType == XmlNodeType.Text)
                    {
                        switch (element)
                        {
                            case "crmorganizationservice": //Display the text in each element.
                                OrganizationUri = new Uri(reader.Value);
                                break;
                            case "username": //Display the end of the element.
                                Credentials.UserName.UserName = reader.Value;
                                break;
                            case "password": //Display the end of the element.
                                Credentials.UserName.Password = reader.Value;
                                break;
                        }

                    }
                }

                reader.Close();

                _serviceProxy = new Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy(OrganizationUri, null, Credentials, DeviceCrednetials);
                _serviceProxy.EnableProxyTypes();

                _service = (Microsoft.Xrm.Sdk.IOrganizationService)_serviceProxy;
            }
            catch(Exception ex)
            {
                string exeption = ex.Message;
            }
        }
        #endregion Private Methods

        

        

        
    }
}
