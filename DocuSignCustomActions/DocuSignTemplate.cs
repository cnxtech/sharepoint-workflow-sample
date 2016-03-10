using System;
using System.Collections.Generic;
using Newtonsoft.Json;

using DocuSign.eSign.Api;
using DocuSign.eSign.Client;
using DocuSign.eSign.Model;

namespace DocuSignCustomActions
{
    /// <summary>
    /// Used to send from DocuSign templates
    /// </summary>
    public static class DocuSignTemplate
    {
        /// <summary>
        /// Send from template that has a fixed role (Person)
        /// </summary>
        /// <param name="DSEnvironment">url to DocuSign</param>
        /// <param name="DSUsername">DocuSign email</param>
        /// <param name="DSPassword">DocuSign password</param>
        /// <param name="DSIntegratorKey">DocuSign IK</param>
        /// <param name="DSTemplateId">DocuSign TemplateId</param>
        /// <param name="NumberOfPackages">Total Number of Packages</param>
        public static string SendFixedRoleTemplateFromWorkflow(string DSEnvironment, string DSUsername, string DSPassword, string DSIntegratorKey, string DSTemplateId, Dictionary<string, string> values)
        {
            // instantiate a new api client
            string APIEndPoint = string.Format("https://{0}/restapi", DSEnvironment);
            ApiClient apiClient = new ApiClient(APIEndPoint);
            Configuration.Default.ApiClient = apiClient;

            // create inline JSON auth header
            string authHeader = "{\"Username\":\"" + DSUsername + "\", \"Password\":\"" + DSPassword + "\", \"IntegratorKey\":\"" + DSIntegratorKey + "\"}";
            Configuration.Default.DefaultHeader.Clear();
            Configuration.Default.AddDefaultHeader("X-DocuSign-Authentication", authHeader);

            // we will retrieve this from the login() call
            string accountId = null;

            //===========================================================
            // Step 1: Login API
            //===========================================================

            // the authentication api uses the apiClient (and X-DocuSign-Authentication header) that are set in Configuration object
            AuthenticationApi authApi = new AuthenticationApi();
            LoginInformation loginInfo = authApi.Login();

            // find the default account for this user
            foreach (LoginAccount loginAcct in loginInfo.LoginAccounts)
            {
                if (loginAcct.IsDefault == "true")
                {
                    accountId = loginAcct.AccountId;
                    break;
                }
            }
            if (accountId == null)
            {
                // if no default found set to first account
                accountId = loginInfo.LoginAccounts[0].AccountId;
            }

            //===========================================================
            // Step 2: Get Template API
            //===========================================================

            var envTemp = new EnvelopeTemplate();
            var templatesApi = new TemplatesApi();
            // get all template information from DocuSign based on the templateId
            envTemp = templatesApi.Get(accountId, DSTemplateId);

            // we will parse the template definition that's returned and read all the recipient
            // roles of type |signer|. DocuSign supports an additional 6 types of recipients.
            var templateRecips = envTemp.Recipients;
            var signers = new List<Signer>();
            signers = templateRecips.Signers;

            //===========================================================
            // Step 3: Create Envelope API
            //===========================================================

            var templateRolesList = new List<TemplateRole>();
            UpdateRecipientTabsWithValues(values, signers, templateRolesList);

            // create a new envelope definition
            var envDef = new EnvelopeDefinition();
            envDef.EmailSubject = "[Sent from Sharepoint Workflow] - Please sign this doc:";

            // use the template that is passed in to this function
            envDef.TemplateId = DSTemplateId;

            // assign the the roles information we read from the template is step 2
            envDef.TemplateRoles = templateRolesList;

            // set envelope status to "sent" to immediately send the signature request
            envDef.Status = "sent";

            // |EnvelopesApi| contains methods related to creating and sending Envelopes (aka signature requests)
            EnvelopesApi envelopesApi = new EnvelopesApi();
            EnvelopeSummary envelopeSummary = envelopesApi.CreateEnvelope(accountId, envDef);

            return envelopeSummary.EnvelopeId;
        } 

        /// <summary>
        /// Update the Total Number of Packages tab with the value
        /// </summary>
        /// <param name="values">a pair of name/value objects to update</param>
        /// <param name="signers">Collection of signer objects</param>
        /// <param name="templateRolesList">Roles with tab information for template</param>
        private static void UpdateRecipientTabsWithValues(Dictionary<string, string> values, List<Signer> signers, List<TemplateRole> templateRolesList)
        {
            foreach (var signer in signers)
            {
                var tRole = new TemplateRole { RoleName = signer.RoleName, Name = signer.Name, Email = signer.Email };
                // read the tabs for each signer and look for a match to the key we're updating
                // so we can pre-populate it's value in the document
                var tabs = signer.Tabs;
                if (tabs != null)
                {
                    foreach (string val in values.Keys)
                    {
                        // check to see if any text tabs assigned to this recipient match any of our values
                        var textTabs = new List<Text>();
                        if (tabs.TextTabs != null)
                        {
                            textTabs = tabs.TextTabs;
                            for (int j = 0; j < textTabs.Count; j++)
                            {
                                // if we find a textTab with a matching name - update its value
                                if (string.Compare(textTabs[j].TabLabel, val) == 0)
                                {
                                    tabs.TextTabs[j].Value = values[val];
                                }
                            }
                        }
                    }
                }
                // assign the tabs to the template role and add to roles list
                tRole.Tabs = tabs;
                templateRolesList.Add(tRole);
            }
        }
    }
}
