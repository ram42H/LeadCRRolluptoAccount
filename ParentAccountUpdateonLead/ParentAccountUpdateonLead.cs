// =====================================================================
//  This file is part of the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================

//<snippetAutoRouteLead>
using System;
using System.Activities;
using System.Collections.ObjectModel;

using Microsoft.Crm.Sdk.Messages;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace Futuredontics.ParentAccountUpdateonLead
{
    public sealed class ParentAccountUpdateonLead : CodeActivity
    {
        /// <summary>
        /// This method first retrieves the lead. Afterwards, it checks the Parent Account id
        /// If Existing Account has data , all the Campaign responses of Lead are rolled upto Account else we remove the Cr's from Account that are related to Lead
        /// </summary>


        protected override void Execute(CodeActivityContext executionContext)
        {

            #region Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            #endregion

            #region Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            #region Retrieve the lead GUID
            Guid leadId = context.PrimaryEntityId;
            tracingService.Trace("Lead Guid -" + leadId);
            #endregion

            Entity campaignresponse = new Entity("campaignresponse");
            EntityReference acc = new EntityReference("account");

            #region Get Parent Account Id using Query by Attribute
            QueryByAttribute AccountQueryByleadId = new QueryByAttribute("lead");
            AccountQueryByleadId.AddAttributeValue("leadid", leadId);
            AccountQueryByleadId.ColumnSet = new ColumnSet("leadid", "parentaccountid");
            EntityCollection leadrecords = service.RetrieveMultiple(AccountQueryByleadId);

            foreach (Entity leadrecord in leadrecords.Entities)
            {
                Guid cr = Guid.Empty;

                string fetchLeadCR = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                                                "<entity name='campaignresponse'>" +
                                                  "<attribute name='subject' />" +
                                                  "<attribute name='activityid' />" +
                                                  "<attribute name='regardingobjectid' />" +
                                                  "<attribute name='responsecode' />" +
                                                  "<attribute name='customer' />" +
                                                  "<attribute name='statecode' />" +
                                                  "<attribute name='statuscode' />" +
                                                  "<order attribute='subject' descending='false' />" +
                                                  "<filter type='and'>" +
                                                    "<condition attribute='fdx_reconversionlead' operator='eq'  value='" + leadId + "' />" +
                                                  "</filter>" +
                                                "</entity>" +
                                              "</fetch>";

                EntityCollection LeadCrs = service.RetrieveMultiple(new FetchExpression(fetchLeadCR));
                if (LeadCrs.Entities.Count > 0)
                {
                    foreach (Entity leadcr in LeadCrs.Entities)
                    {
                        cr = ((Guid)leadcr["activityid"]);

                        tracingService.Trace("CR Guid-" + cr);

                        if (leadrecord.Attributes.Contains("parentaccountid"))
                        {
                            acc.Id = ((EntityReference)leadrecord["parentaccountid"]).Id;
                            acc.Name = ((EntityReference)leadrecord["parentaccountid"]).Name;

                            tracingService.Trace("Account Guid-" + acc.Id);
                            tracingService.Trace("Account Name -" + acc.Name);



                            EntityReference LeadAccount = new EntityReference("account", acc.Id);

                            Entity customer = new Entity("activityparty");
                            customer.Attributes["partyid"] = LeadAccount;
                            EntityCollection Customerentity = new EntityCollection();
                            Customerentity.Entities.Add(customer);

                            campaignresponse["customer"] = Customerentity;

                            tracingService.Trace("Customer -" + Customerentity);

                            campaignresponse["activityid"] = cr;

                            service.Update(campaignresponse);

                            tracingService.Trace("CR is updated with Account on Customer");
                        }
                        else
                        {
                            campaignresponse["customer"] = null;

                            campaignresponse["activityid"] = cr;

                            service.Update(campaignresponse);

                            tracingService.Trace("Account is cleared on Customer-Campaign Response");
                        }
                    }
                }
            #endregion
            }
        }
    }
}
