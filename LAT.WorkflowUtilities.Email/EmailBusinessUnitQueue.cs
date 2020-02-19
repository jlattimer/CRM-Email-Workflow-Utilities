using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace LAT.WorkflowUtilities.Email
{
    public class EmailBusinessUnitQueue : WorkFlowActivityBase
    {
        public EmailBusinessUnitQueue() : base(typeof(EmailBusinessUnitQueue)) { }

        [RequiredArgument]
        [Input("Email To Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToSend { get; set; }

        [RequiredArgument]
        [Input("Recipient Business Unit")]
        [ReferenceTarget("businessunit")]
        public InArgument<EntityReference> RecipientBusinessUnit { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Send Email?")]
        public InArgument<bool> SendEmail { get; set; }

        [Output("Email Sent")]
        public OutArgument<bool> EmailSent { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            EntityReference emailToSend = EmailToSend.Get(context);
            EntityReference recipientBusinessUnit = RecipientBusinessUnit.Get(context);
            bool sendEmail = SendEmail.Get(context);

            List<Entity> toList = new List<Entity>();

            Entity email = RetrieveEmail(localContext.OrganizationService, emailToSend.Id);

            if (email == null)
            {
                EmailSent.Set(context, false);
                return;
            }

            //Add any pre-defined recipients specified to the array               
            foreach (Entity activityParty in email.GetAttributeValue<EntityCollection>("to").Entities)
            {
                toList.Add(activityParty);
            }

            List<EntityReference> buQueues = GetBuQueues(localContext.OrganizationService, recipientBusinessUnit.Id);

            if (buQueues.Count < 1)
            {
                EmailSent.Set(context, false);
                localContext.Trace("GetBuQueue found no default queues for the Business Unit.");
                return;
            }
            else if (buQueues.Count > 1)
            {
                localContext.Trace("GetBuQueue found more than one default queue for the Business Unit.");
            }

            toList = ProcessQueues(buQueues, toList);

            //Update the email
            email["to"] = toList.ToArray();
            localContext.OrganizationService.Update(email);

            //Send
            if (sendEmail)
            {
                SendEmailRequest request = new SendEmailRequest
                {
                    EmailId = emailToSend.Id,
                    TrackingToken = string.Empty,
                    IssueSend = true
                };

                localContext.OrganizationService.Execute(request);
                EmailSent.Set(context, true);
            }
            else
            {
                EmailSent.Set(context, false);
            }
        }

        private static Entity RetrieveEmail(IOrganizationService service, Guid emailId)
        {
            return service.Retrieve("email", emailId, new ColumnSet("to"));
        }

        private static List<Entity> ProcessQueues(List<EntityReference> queues, List<Entity> toList)
        {
            foreach (EntityReference e in queues)
            {
                Entity activityParty =
                    new Entity("activityparty")
                    {
                        ["partyid"] = e
                    };

                if (toList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.Id)) continue;

                toList.Add(activityParty);
            }

            return toList;
        }

        private static List<EntityReference> GetBuQueues(IOrganizationService service, Guid businessUnitId)
        {
            //Query for the business unit members
            QueryExpression query = new QueryExpression
            {
                EntityName = "team",
                ColumnSet = new ColumnSet("queueid"),
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "team",
                        LinkFromAttributeName = "businessunitid",
                        LinkToEntityName = "businessunit",
                        LinkToAttributeName = "businessunitid",
                        LinkCriteria = new FilterExpression
                        {
                            FilterOperator = LogicalOperator.And,
                            Conditions =
                            {
                                new ConditionExpression
                                {
                                    AttributeName = "businessunitid",
                                    Operator = ConditionOperator.Equal,
                                    Values = { businessUnitId }
                                }
                            }
                        }
                    }
                },
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression
                        {
                            AttributeName = "isdefault",
                            Operator = ConditionOperator.Equal,
                            Values = { false }
                        }
                    }
                }
            };

            EntityCollection results = service.RetrieveMultiple(query);
            List<EntityReference> queues = new List<EntityReference>();

            foreach (Entity t in results.Entities)
            {
                queues.Add((EntityReference)t["queueid"]);
            }

            return queues;
        }
    }
}