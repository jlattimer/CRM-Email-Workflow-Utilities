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

            EntityCollection buQueues = GetBuQueues(localContext.OrganizationService, recipientBusinessUnit.Id);

            if (buQueues.Entities.Count < 1)
            {
                localContext.Trace("GetBuQueue found no default queues for the Business Unit.");
            }
            else
            {
                if (buQueues.Entities.Count > 1)
                {
                    localContext.Trace("GetBuQueue found more than one default queue for the Business Unit.");
                }

                toList = ProcessQueues(buQueues, toList);

                //Update the email
                email["to"] = toList.ToArray();
                localContext.OrganizationService.Update(email);
            }

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

        private static List<Entity> ProcessQueues(EntityCollection queues, List<Entity> toList)
        {
            foreach (Entity e in queues.Entities)
            {
                Entity activityParty =
                    new Entity("activityparty")
                    {
                        ["partyid"] = new EntityReference(e.LogicalName, e.Id)
                    };

                if (toList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.Id)) continue;

                toList.Add(activityParty);
            }

            return toList;
        }

        private static EntityCollection GetBuQueues(IOrganizationService service, Guid businessUnitId)
        {
            FilterExpression filter = new FilterExpression();
            filter.FilterOperator = LogicalOperator.And;
            filter.Conditions.Add(new ConditionExpression
            {
                AttributeName = "businessunitid",
                Operator = ConditionOperator.Equal,
                Values = { businessUnitId }
            });
            filter.Conditions.Add(new ConditionExpression
            {
                AttributeName = "isdefault",
                Operator = ConditionOperator.Equal,
                Values = { true }
            });

            QueryExpression query = new QueryExpression
            {
                EntityName = "queue",
                ColumnSet = new ColumnSet(false),
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "queue",
                        LinkFromAttributeName = "queueid",
                        LinkToEntityName = "team",
                        LinkToAttributeName = "queueid",
                        LinkCriteria = filter
                    }
                }
            };

            return service.RetrieveMultiple(query);
        }
    }
}