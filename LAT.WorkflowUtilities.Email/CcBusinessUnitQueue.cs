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
    public class CcBusinessUnitQueue : WorkFlowActivityBase
    {
        public CcBusinessUnitQueue() : base(typeof(EmailBusinessUnitQueue)) { }

        [RequiredArgument]
        [Input("Email To Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToSend { get; set; }

        [RequiredArgument]
        [Input("CC Business Unit")]
        [ReferenceTarget("businessunit")]
        public InArgument<EntityReference> RecipientBusinessUnit { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Send Email?")]
        public InArgument<bool> SendEmail { get; set; }

        [Output("Email Sent")]
        public OutArgument<bool> EmailSent { get; set; }

        [Output("Users Added")]
        public OutArgument<int> UsersAdded { get; set; }

        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            EntityReference emailToSend = EmailToSend.Get(context);
            EntityReference recipientBusinessUnit = RecipientBusinessUnit.Get(context);
            bool sendEmail = SendEmail.Get(context);

            List<Entity> ccList = new List<Entity>();

            Entity email = RetrieveEmail(localContext.OrganizationService, emailToSend.Id);

            if (email == null)
            {
                EmailSent.Set(context, false);
                UsersAdded.Set(context, 0);
                return;
            }

            //Add any pre-defined recipients specified to the array               
            foreach (Entity activityParty in email.GetAttributeValue<EntityCollection>("cc").Entities)
            {
                ccList.Add(activityParty);
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

                ccList = ProcessQueues(buQueues, ccList);

                //Update the email
                email["cc"] = ccList.ToArray();
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

            UsersAdded.Set(context, ccList.Count);
        }

        private static Entity RetrieveEmail(IOrganizationService service, Guid emailId)
        {
            return service.Retrieve("email", emailId, new ColumnSet("cc"));
        }

        private static List<Entity> ProcessQueues(EntityCollection queues, List<Entity> ccList)
        {
            foreach (Entity e in queues.Entities)
            {
                Entity activityParty =
                    new Entity("activityparty")
                    {
                        ["partyid"] = new EntityReference(e.LogicalName, e.Id)
                    };

                if (ccList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.Id)) continue;

                ccList.Add(activityParty);
            }

            return ccList;
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
                },
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression
                        {
                            AttributeName = "statecode",
                            Operator = ConditionOperator.Equal,
                            Values = { 0 }
                        }
                    }
                }
            };

            return service.RetrieveMultiple(query);
        }
    }
}