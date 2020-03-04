using FakeXrmEasy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class EmailBusinessUnitQueueTests
    {
        #region Test Initialization and Cleanup
        // Use ClassInitialize to run code before running the first test in the class
        [ClassInitialize()]
        public static void ClassInitialize(TestContext testContext) { }

        // Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void ClassCleanup() { }

        // Use TestInitialize to run code before running each test 
        [TestInitialize()]
        public void TestMethodInitialize() { }

        // Use TestCleanup to run code after each test has run
        [TestCleanup()]
        public void TestMethodCleanup() { }
        #endregion

        [TestMethod]
        public void EmailBusinessUnitQueue_Business_Unit_With_No_Existing_Default_Team()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team.Id,
                ["owningteam"] = team.Id,
                ["businessunitid"] = businessUnit.Id
            };

            // Relationship from team to its queue
            team["queueid"] = queue.Id;

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, team, systemUser1, teammembership, queue });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_Business_Unit_With_Default_Team_But_No_Queues()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, team, systemUser1, teammembership });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_Business_Unit_With_1_Default_Team_And_Queue()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team.Id,
                ["owningteam"] = team.Id,
                ["businessunitid"] = businessUnit.Id,
            };

            // Relationship from team to its queue
            team["queueid"] = queue.Id;

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, team, systemUser1, teammembership, queue });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_Business_Unit_With_1_Default_Team_And_Queue_Existing_To()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["cc"] = new EntityCollection()
            };

            // Existing recipient
            Guid id2 = Guid.NewGuid();
            Entity activityParty = new Entity("activityparty")
            {
                Id = id2,
                ["activitypartyid"] = id2,
                ["activityid"] = new EntityReference("email", id),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityParty);
            email["to"] = to;

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team.Id,
                ["owningteam"] = team.Id,
                ["businessunitid"] = businessUnit.Id
            };

            // Relationship from team to its queue
            team["queueid"] = queue.Id;

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, team, systemUser1, teammembership, queue, activityParty });

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_Business_Unit_With_1_Default_Team_And_Queue_Existing_CC()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            // Existing recipient
            Guid id2 = Guid.NewGuid();
            Entity activityParty = new Entity("activityparty")
            {
                Id = id2,
                ["activitypartyid"] = id2,
                ["activityid"] = new EntityReference("email", id),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityParty);
            email["cc"] = to;

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team.Id,
                ["owningteam"] = team.Id,
                ["businessunitid"] = businessUnit.Id
            };

            // Relationship from team to its queue
            team["queueid"] = queue.Id;

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, team, systemUser1, teammembership, queue, activityParty });

            const int expected = 2;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<CcBusinessUnitQueue>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_Business_Unit_With_1_Default_Team_And_2_Queues()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue1 = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team.Id,
                ["owningteam"] = team.Id,
                ["businessunitid"] = businessUnit.Id
            };

            // Relationship from team to its queue
            team["queueid"] = queue1.Id;

            // Second queue, owned by the tean but not the default queue
            Entity queue2 = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team.Id,
                ["owningteam"] = team.Id,
                ["businessunitid"] = businessUnit.Id
            };

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, team, systemUser1, teammembership, queue1, queue2 });

            const int expected = 1;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_Business_Unit_With_2_Teams_And_2_Queues_Each()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team1 = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership1 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team1.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue1 = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team1.Id,
                ["owningteam"] = team1.Id,
                ["businessunitid"] = businessUnit.Id
            };

            // Relationship from team to its queue
            team1["queueid"] = queue1.Id;

            // Second team can't be the default
            Entity team2 = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership2 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team2.Id,
                ["systemuserid"] = systemUser2.Id
            };

            Entity queue2 = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team2.Id,
                ["owningteam"] = team2.Id,
                ["businessunitid"] = businessUnit.Id
            };

            // Relationship from team to its queue
            team2["queueid"] = queue2.Id;

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, team1, systemUser1, teammembership1, queue1, team2, systemUser2, teammembership2, queue2 });

            const int expected = 1; // Just the default team should be returned

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_Business_Unit_With_Disabled_Queue()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid id = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = id,
                ["activityid"] = id,
                ["to"] = new EntityCollection()
            };

            Entity businessUnit = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit.Id
            };

            Entity teammembership = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team.Id,
                ["owningteam"] = team.Id,
                ["businessunitid"] = businessUnit.Id,
                ["statecode"] = new OptionSetValue(1),
                ["statuscode"] = new OptionSetValue(2)
            };

            // Relationship from team to its queue
            team["queueid"] = queue.Id;

            var inputs = new Dictionary<string, object>
            {
                { "EmailToSend", email.ToEntityReference() },
                { "RecipientBusinessUnit", businessUnit.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit, team, systemUser1, teammembership, queue });

            const int expected = 0;

            //Act
            var result = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputs);

            //Assert
            Assert.AreEqual(expected, result["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_1_Business_Unit_To_And_CC_With_Default_Teams_And_Queues()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid idEmail = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = idEmail,
                ["activityid"] = idEmail,
                ["cc"] = new EntityCollection()
            };

            // Existing recipient for To field
            Guid idTo = Guid.NewGuid();
            Entity activityPartyTo = new Entity("activityparty")
            {
                Id = idTo,
                ["activitypartyid"] = idTo,
                ["activityid"] = new EntityReference("email", idEmail),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityPartyTo);
            email["to"] = to;

            // Existing recipients for Cc field
            Guid idCc1 = Guid.NewGuid();
            Entity activityPartyCc1 = new Entity("activityparty")
            {
                Id = idCc1,
                ["activitypartyid"] = idCc1,
                ["activityid"] = new EntityReference("email", idEmail),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };
            Guid idCc2 = Guid.NewGuid();
            Entity activityPartyCc2 = new Entity("activityparty")
            {
                Id = idCc1,
                ["activitypartyid"] = idCc2,
                ["activityid"] = new EntityReference("email", idEmail),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection cc = new EntityCollection();
            cc.Entities.Add(activityPartyCc1);
            cc.Entities.Add(activityPartyCc2);
            email["cc"] = cc;

            Entity businessUnit1 = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team1 = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit1.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit1.Id
            };

            Entity teammembership1 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team1.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue1 = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team1.Id,
                ["owningteam"] = team1.Id,
                ["businessunitid"] = businessUnit1.Id
            };

            // Relationship from team to its queue
            team1["queueid"] = queue1.Id;

            EntityReference idInputEmail = email.ToEntityReference();
            var inputsFirst = new Dictionary<string, object>
            {
                { "EmailToSend", idInputEmail },
                { "RecipientBusinessUnit", businessUnit1.ToEntityReference() },
                { "SendEmail", false }
            };

            var inputsSecond = new Dictionary<string, object>
            {
                { "EmailToSend", idInputEmail },
                { "RecipientBusinessUnit", businessUnit1.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit1, team1, systemUser1, teammembership1, queue1, activityPartyTo, activityPartyCc1 });

            const int expectedAfterTo = 2;
            const int expectedAfterCc = 3;

            //Act
            var resultFirstAct = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputsFirst);

            //Assert
            Assert.AreEqual(expectedAfterTo, resultFirstAct["UsersAdded"]);

            //Second Act
            var resultSecondAct = xrmFakedContext.ExecuteCodeActivity<CcBusinessUnitQueue>(workflowContext, inputsSecond);

            //Assert
            Assert.AreEqual(expectedAfterCc, resultSecondAct["UsersAdded"]);
        }

        [TestMethod]
        public void EmailBusinessUnitQueue_1_Business_Unit_To_1_Separate_BU_CC_With_Default_Teams_And_Queues()
        {
            //Arrange
            XrmFakedWorkflowContext workflowContext = new XrmFakedWorkflowContext();

            Guid idEmail = Guid.NewGuid();
            Entity email = new Entity("email")
            {
                Id = idEmail,
                ["activityid"] = idEmail,
                ["cc"] = new EntityCollection()
            };

            // Existing recipient for To field
            Guid idTo = Guid.NewGuid();
            Entity activityPartyTo = new Entity("activityparty")
            {
                Id = idTo,
                ["activitypartyid"] = idTo,
                ["activityid"] = new EntityReference("email", idEmail),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection to = new EntityCollection();
            to.Entities.Add(activityPartyTo);
            email["to"] = to;

            // Existing recipients for Cc field
            Guid idCc1 = Guid.NewGuid();
            Entity activityPartyCc1 = new Entity("activityparty")
            {
                Id = idCc1,
                ["activitypartyid"] = idCc1,
                ["activityid"] = new EntityReference("email", idEmail),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };
            Guid idCc2 = Guid.NewGuid();
            Entity activityPartyCc2 = new Entity("activityparty")
            {
                Id = idCc1,
                ["activitypartyid"] = idCc2,
                ["activityid"] = new EntityReference("email", idEmail),
                ["partyid"] = new EntityReference("contact", Guid.NewGuid()),
                ["participationtypemask"] = new OptionSetValue(2)
            };

            EntityCollection cc = new EntityCollection();
            cc.Entities.Add(activityPartyCc1);
            cc.Entities.Add(activityPartyCc2);
            email["cc"] = cc;

            // First BU
            Entity businessUnit1 = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team1 = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit1.Id
            };

            Entity systemUser1 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit1.Id
            };

            Entity teammembership1 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team1.Id,
                ["systemuserid"] = systemUser1.Id
            };

            Entity queue1 = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team1.Id,
                ["owningteam"] = team1.Id,
                ["businessunitid"] = businessUnit1.Id
            };

            // Relationship from team to its queue
            team1["queueid"] = queue1.Id;

            // Second BU
            Entity businessUnit2 = new Entity("businessunit")
            {
                Id = Guid.NewGuid()
            };

            Entity team2 = new Entity("team")
            {
                Id = Guid.NewGuid(),
                ["isdefault"] = true,
                ["businessunitid"] = businessUnit2.Id
            };

            Entity systemUser2 = new Entity("systemuser")
            {
                Id = Guid.NewGuid(),
                ["internalemailaddress"] = null,
                ["isdisabled"] = false,
                ["businessunitid"] = businessUnit2.Id
            };

            Entity teammembership2 = new Entity("teammembership")
            {
                Id = Guid.NewGuid(),
                ["teamid"] = team2.Id,
                ["systemuserid"] = systemUser2.Id
            };

            Entity queue2 = new Entity("queue")
            {
                Id = Guid.NewGuid(),
                ["ownerid"] = team2.Id,
                ["owningteam"] = team2.Id,
                ["businessunitid"] = businessUnit2.Id
            };

            // Relationship from team to its queue
            team2["queueid"] = queue2.Id;

            EntityReference idInputEmail = email.ToEntityReference();
            var inputsFirst = new Dictionary<string, object>
            {
                { "EmailToSend", idInputEmail },
                { "RecipientBusinessUnit", businessUnit1.ToEntityReference() },
                { "SendEmail", false }
            };

            var inputsSecond = new Dictionary<string, object>
            {
                { "EmailToSend", idInputEmail },
                { "RecipientBusinessUnit", businessUnit2.ToEntityReference() },
                { "SendEmail", false }
            };

            XrmFakedContext xrmFakedContext = new XrmFakedContext();
            xrmFakedContext.Initialize(new List<Entity> { email, businessUnit1, team1, systemUser1, teammembership1, queue1, businessUnit2, team2, systemUser2, teammembership2, queue2, activityPartyTo, activityPartyCc1 });

            const int expectedAfterTo = 2;
            const int expectedAfterCc = 3;

            //Act
            var resultFirstAct = xrmFakedContext.ExecuteCodeActivity<EmailBusinessUnitQueue>(workflowContext, inputsFirst);

            //Assert
            Assert.AreEqual(expectedAfterTo, resultFirstAct["UsersAdded"]);

            //Second Act
            var resultSecondAct = xrmFakedContext.ExecuteCodeActivity<CcBusinessUnitQueue>(workflowContext, inputsSecond);

            //Assert
            Assert.AreEqual(expectedAfterCc, resultSecondAct["UsersAdded"]);
        }
    }
}