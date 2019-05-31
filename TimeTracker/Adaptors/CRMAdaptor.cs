using System;
using System.Collections.Generic;
using System.ServiceModel.Description;
using System.Text.RegularExpressions;
using System.Xml;
using TimeTracker.Common;
using TimeTracker.Entity;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk.WebServiceClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace TimeTracker.Adaptors
{
    public class CRMAdaptor
    {
        public Guid license;
        ClientCredentials Credentials;
        IOrganizationService _service;

        #region Public Methods
        public CRMAdaptor()
        {
            
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
                throw ex;
            }
        }

        public bool? DeactivateExternalCommentInCRM(TimeItem timeItem)
        {
            try
            {
                Microsoft.Xrm.Sdk.Entity comment = _service.Retrieve("hsal_externalcomments", (Guid)timeItem.crmExternalCommentId, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));

                Microsoft.Crm.Sdk.Messages.SetStateRequest req = new Microsoft.Crm.Sdk.Messages.SetStateRequest();
                req.EntityMoniker = new Microsoft.Xrm.Sdk.EntityReference("hsal_externalcomments", comment.Id);
                req.State = new Microsoft.Xrm.Sdk.OptionSetValue(2);
                req.Status = new Microsoft.Xrm.Sdk.OptionSetValue((-1));
                _service.Execute(req);

                return true;
            }
            catch (Exception ex)
            {
                //log error
                throw ex;
            }
        }

        public bool? UpdateExternalCommentInCRM(string ticketnumber, TimeItem timeItem)
        {
            try
            {
                Microsoft.Xrm.Sdk.Entity comment = _service.Retrieve("hsal_externalcomments", (Guid)timeItem.crmExternalCommentId, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));

                Microsoft.Crm.Sdk.Messages.SetStateRequest req = new Microsoft.Crm.Sdk.Messages.SetStateRequest();
                req.EntityMoniker = new Microsoft.Xrm.Sdk.EntityReference("hsal_externalcomments", comment.Id);
                req.State = new Microsoft.Xrm.Sdk.OptionSetValue(0);
                req.Status = new Microsoft.Xrm.Sdk.OptionSetValue((-1));
                _service.Execute(req);

                if (timeItem.title != null)
                {
                    comment["subject"] = timeItem.title.Replace(ticketnumber, "");
                }

                if (timeItem.description != null)
                {
                    comment["description"] = timeItem.description;
                }

                _service.Update(comment);

                req.EntityMoniker = new Microsoft.Xrm.Sdk.EntityReference("hsal_externalcomments", comment.Id);
                req.State = new Microsoft.Xrm.Sdk.OptionSetValue(1);
                req.Status = new Microsoft.Xrm.Sdk.OptionSetValue((2));
                _service.Execute(req);

                return true;
            }
            catch (Exception ex)
            {
                //log error
                throw ex;
            }
        }

        public Guid? CreateExternalComment(string ticketnumber, TimeItem timeItem)
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

                Microsoft.Xrm.Sdk.Entity comment = new Microsoft.Xrm.Sdk.Entity("hsal_externalcomments");

                if (timeItem.title != null)
                {
                    comment["subject"] = timeItem.title.Replace(ticketnumber, "");
                }

                if (timeItem.description != null)
                {
                    comment["description"] = timeItem.description;
                }

                comment["regardingobjectid"] = CaseResults.Entities[0].ToEntityReference();

                Guid externalCommentId = _service.Create(comment);

                Microsoft.Crm.Sdk.Messages.SetStateRequest req = new Microsoft.Crm.Sdk.Messages.SetStateRequest();
                req.EntityMoniker = new Microsoft.Xrm.Sdk.EntityReference("hsal_externalcomments", externalCommentId);
                req.State = new Microsoft.Xrm.Sdk.OptionSetValue(1);
                req.Status = new Microsoft.Xrm.Sdk.OptionSetValue((2));
                _service.Execute(req);

                return externalCommentId;
            }
            catch (Exception ex)
            {
                //log error
                throw ex;
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
                throw ex;
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
                try
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

                                item.isCRMSubmitted = item.crmTaskId != null ? true : false;
                            }

                            if (item.crmExternalCommentId != null)
                            {
                                if (item.isExternalComment == true)
                                {
                                    UpdateExternalCommentInCRM(word, item);
                                }
                                else
                                {
                                    DeactivateExternalCommentInCRM(item);
                                }
                            }
                            else
                            {
                                if (item.isExternalComment == true)
                                {
                                    item.crmExternalCommentId = CreateExternalComment(word, item);
                                }
                            }
                        }
                    }

                }
                catch(Exception ex)
                {
                    throw ex;
                }
            }
            return timeItems;
        }

        #endregion Public Methods

        #region Private Methods
        public bool GetCRMConnection()
        {
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
                            case "userlicence": //Display the text in each element.
                                bool isValid = Guid.TryParse(reader.Value, out license);
                                if (!isValid)
                                {
                                    MessageBox.Show("The license is in an invalid format.  Time Tracker will not be able to sycn to JARVIS");
                                    return false;
                                }
                                break;
                        }

                    }
                }

                reader.Close();

                string organizationUrl = "https://csp.crm.dynamics.com";
                string resourceURL = "https://csp.api.crm.dynamics.com" + "/api/data/";
                string clientId = "c4e4407b-66d1-4452-9b05-db0a0ce9baef"; // Client Id
                string appKey = "Sy[?Mk106C2OvHXZ:Krytwj=_XN_KKlh"; //Client Secret

                //Create the Client credentials to pass for authentication
                Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential clientcred = new Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential(clientId, appKey);

                //get the authentication parameters
                Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationParameters authParam = Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationParameters.CreateFromResourceUrlAsync(new Uri(resourceURL)).Result;

                //Generate the authentication context - this is the azure login url specific to the tenant
                string authority = authParam.Authority;

                //request token
                AuthenticationResult authenticationResult = new AuthenticationContext(authority).AcquireTokenAsync(organizationUrl, clientcred).Result;

                //get the token              
                string token = authenticationResult.AccessToken;

                Uri serviceUrl = new Uri(organizationUrl + @"/xrmservices/2011/organization.svc/web?SdkClientVersion=9.1");
                OrganizationWebProxyClient sdkService;

                sdkService = new OrganizationWebProxyClient(serviceUrl, false);
                sdkService.CallerId = license;
                sdkService.HeaderToken = token;

                _service = (Microsoft.Xrm.Sdk.IOrganizationService)sdkService != null ? (Microsoft.Xrm.Sdk.IOrganizationService)sdkService : null;
                Microsoft.Xrm.Sdk.Entity user = _service.Retrieve("systemuser", license, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                return true;

            }
            catch(Exception ex)
            {
                MessageBox.Show("You do not have a valid license for JARIVS and will not be able to sync your time");
                return false;
            }
        }
        #endregion Private Methods
    }
}
