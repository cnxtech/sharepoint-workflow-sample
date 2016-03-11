using Microsoft.SharePoint;
using Microsoft.SharePoint.WorkflowActions;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Workflow.ComponentModel;
using System.Workflow.ComponentModel.Compiler;

namespace DocuSignCustomActions
{
    /// <summary>
    /// Activity/Custom Action for sending a DocuSign Template
    /// </summary>
    public class SendTemplateAction : Activity
    {
        public static DependencyProperty __ContextProperty = DependencyProperty.Register("__Context", typeof(WorkflowContext), typeof(SendTemplateAction));
        
        [Description("The site context")]
        [Category("User")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public WorkflowContext __Context
        {
            get
            {
                return ((WorkflowContext)(base.GetValue(SendTemplateAction.__ContextProperty)));
            }
            set
            {
                base.SetValue(SendTemplateAction.__ContextProperty, value);
            }
        }

        public static DependencyProperty __ListIdProperty = DependencyProperty.Register("__ListId", typeof(string), typeof(SendTemplateAction));

        [ValidationOption(ValidationOption.Required)]
        public string __ListId
        {
            get
            {
                return ((string)(base.GetValue(SendTemplateAction.__ListIdProperty)));
            }
            set
            {
                base.SetValue(SendTemplateAction.__ListIdProperty, value);
            }
        }

        public static DependencyProperty __ListItemProperty = DependencyProperty.Register("__ListItem", typeof(int), typeof(SendTemplateAction));

        [ValidationOption(ValidationOption.Required)]
        public int __ListItem
        {
            get
            {
                return ((int)(base.GetValue(SendTemplateAction.__ListItemProperty)));
            }
            set
            {
                base.SetValue(SendTemplateAction.__ListItemProperty, value);
            }
        }

        public static DependencyProperty __ActivationPropertiesProperty = DependencyProperty.Register("__ActivationProperties", typeof(Microsoft.SharePoint.Workflow.SPWorkflowActivationProperties), typeof(SendTemplateAction));

        [ValidationOption(ValidationOption.Required)]
        public Microsoft.SharePoint.Workflow.SPWorkflowActivationProperties __ActivationProperties
        {
            get
            {
                return (Microsoft.SharePoint.Workflow.SPWorkflowActivationProperties)base.GetValue(SendTemplateAction.__ActivationPropertiesProperty);
            }
            set
            {
                base.SetValue(SendTemplateAction.__ActivationPropertiesProperty, value);
            }
        }



        public static DependencyProperty UserGroupProperty = DependencyProperty.Register("UserGroup", typeof(string), typeof(SendTemplateAction));
        [Category("DocuSign"), Browsable(true)]
        [DesignerSerializationVisibility
          (DesignerSerializationVisibility.Visible)]
        public string UserGroup
        {
            get
            {
                return Convert.ToString(base.GetValue(UserGroupProperty));
            }
            set
            {
                base.SetValue(UserGroupProperty, value);
            }
        }

        public static DependencyProperty DSEnvironmentProperty = DependencyProperty.Register("DSEnvironment", typeof(string), typeof(SendTemplateAction));
        [Category("DocuSign"), Browsable(true)]
        [DesignerSerializationVisibility
          (DesignerSerializationVisibility.Visible)]
        public string DSEnvironment
        {
            get
            {
                return Convert.ToString(base.GetValue(DSEnvironmentProperty));
            }
            set
            {
                base.SetValue(DSEnvironmentProperty, value);
            }
        }

        public static DependencyProperty DSTemplateIdProperty = DependencyProperty.Register("DSTemplateId", typeof(string), typeof(SendTemplateAction));
        [Category("DocuSign"), Browsable(true)]
        [DesignerSerializationVisibility
          (DesignerSerializationVisibility.Visible)]
        public string DSTemplateId
        {
            get
            {
                return Convert.ToString(base.GetValue(DSTemplateIdProperty));
            }
            set
            {
                base.SetValue(DSTemplateIdProperty, value);
            }
        }

        public static DependencyProperty NumberOfPackagesProperty = DependencyProperty.Register("NumberOfPackages", typeof(string), typeof(SendTemplateAction));
        [Category("DocuSign"), Browsable(true)]
        [DesignerSerializationVisibility
          (DesignerSerializationVisibility.Visible)]
        public int NumberOfPackages
        {
            get
            {
                return Convert.ToInt32(base.GetValue(NumberOfPackagesProperty));
            }
            set
            {
                base.SetValue(NumberOfPackagesProperty, value);
            }
        }


        public static DependencyProperty DSIntegratorKeyProperty = DependencyProperty.Register("DSIntegratorKey", typeof(string), typeof(SendTemplateAction));
        [Category("DocuSign"), Browsable(true)]
        [DesignerSerializationVisibility
          (DesignerSerializationVisibility.Visible)]
        public string DSIntegratorKey
        {
            get
            {
                return Convert.ToString(base.GetValue(DSIntegratorKeyProperty));
            }
            set
            {
                base.SetValue(DSIntegratorKeyProperty, value);
            }
        }


        public static DependencyProperty DSUsernameProperty = DependencyProperty.Register("DSUsername", typeof(string), typeof(SendTemplateAction));
        [Category("DocuSign"), Browsable(true)]
        [DesignerSerializationVisibility
          (DesignerSerializationVisibility.Visible)]
        public string DSUsername
        {
            get
            {
                return Convert.ToString(base.GetValue(DSUsernameProperty));
            }
            set
            {
                base.SetValue(DSUsernameProperty, value);
            }
        }

        public static DependencyProperty DSPasswordProperty = DependencyProperty.Register("DSPassword", typeof(string), typeof(SendTemplateAction));
        [Category("DocuSign"), Browsable(true)]
        [DesignerSerializationVisibility
          (DesignerSerializationVisibility.Visible)]
        public string DSPassword
        {
            get
            {
                return Convert.ToString(base.GetValue(DSPasswordProperty));
            }
            set
            {
                base.SetValue(DSPasswordProperty, value);
            }
        }

        public static DependencyProperty ReturnResultProperty = DependencyProperty.Register("ReturnResult", typeof(string), typeof(SendTemplateAction));
        [Category("DocuSign"), Browsable(true)]
        [DesignerSerializationVisibility
          (DesignerSerializationVisibility.Visible)]
        public string ReturnResult
        {
            get
            {
                return (string)base.GetValue(ReturnResultProperty);
            }
            set
            {
                base.SetValue(ReturnResultProperty, value);
            }
        }


        /// <summary>
        /// entry point for custom action
        /// </summary>
        /// <param name="executionContext"></param>
        /// <returns></returns>
        protected override ActivityExecutionStatus Execute(ActivityExecutionContext executionContext)
        {
            this.ReturnResult = "Success";

            try
            {
                // some sanity check for the DSEnvironment so it's a docusign url....
                if (!Regex.IsMatch(this.DSEnvironment, @"^[^\.]*\.docusign\.net$", RegexOptions.IgnoreCase))
                {
                    throw new Exception("DSEnvironment parameter is invalid");
                }

                /// create a collection and ad the number of packages from workflow
                var values = new Dictionary<string, string>();
                values.Add("numPackages", this.NumberOfPackages.ToString());

                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    using (SPSite oSite = new SPSite(__Context.Web.Url))
                    {
                        oSite.AllowUnsafeUpdates = true;
                        // actual DocuSign API work would be inside this static method...
                        this.ReturnResult = DocuSignTemplate.SendFixedRoleTemplateFromWorkflow
                            (this.DSEnvironment, this.DSUsername, this.DSPassword, this.DSIntegratorKey, this.DSTemplateId, values);
                    }
                });
            }
            catch (Exception ex)
            {
                this.ReturnResult = ex.Message;
            }
            
            return ActivityExecutionStatus.Closed;
        }
    }
}
