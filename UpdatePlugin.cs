/*
 Scott Grote
Custom Plugin designed to run with custom entity to streamline changes to contact information.

Anytime a request entity is created this plguin will trigger. 
Code retrieves information from request entity and modifies contact 
entitiy info accordingly. Checks account 
 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Collections;

namespace ContactUpdatePlugin
{
    public class UpdatePlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.  
                Entity reqEntity = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {


                    

                    //contact.Attributes["lastname"]="works";
                    //contact.Attributes["firstname"] = "fine";

                    //Relationship relation = new Relationship("sq_account_contact");
                    /*/
                    var cols = new ColumnSet(new String[] { "ss_name", "ss_patient", "ss_desireddoctor" });
                    var reqEnt = svc.Retrieve("ss_request", new Guid("2daf7b97-72d8-ea11-a813-000d3a31f370"), cols);
                    Console.WriteLine("Request Entity Contact: " + reqEnt.Attributes["ss_name"].ToString() + "_______");
                    EntityReference cRef = (EntityReference)reqEnt.Attributes["ss_patient"];
                    EntityReference aRef = (EntityReference)reqEnt.Attributes["ss_desireddoctor"];

                    Console.WriteLine("Request Entity Contact: " + cRef.Id + " __XX_____");

                    var colsPat = new ColumnSet(new String[] { "parentcustomerid" });
                    var reqCont = svc.Retrieve("contact", cRef.Id, colsPat);
                    var newContact = new Entity("contact");
                    reqCont["parentcustomerid"] = aRef;
                    svc.Update(reqCont);
                    *///


                    //Intializes Contact Enitity to update contact, and entity reference 
                    //
                    Entity contact = new Entity("contact");
                    EntityReference accRef;
                    EntityReference contRef;
                    string desire= "failedattempt";


                    if (reqEntity.Attributes.Contains("ss_patient") && reqEntity.Attributes.ContainsKey("ss_desiredinsurance"))//checks for keys
                    {
                        desire =  reqEntity.Attributes["ss_desiredinsurance"].ToString();
                        contRef = (EntityReference) reqEntity.Attributes["ss_patient"];
                    }
                    else
                    {
                        //accRef = new EntityReference("account", new Guid("bb7a8b88-e5d8-ea11-a813-000d3a31f370"));
                        contRef = new EntityReference("contact", new Guid("043727d4-e5d8-ea11-a813-000d3a31f370"));
                    }


                    //creates a fetchXML to request a list of account names
                    string fetchXml2 = @"<fetch >
                                        <entity name='account' >
                                            <attribute name='name' />
                                            <attribute name='accountid'/>
                                               <order attribute='name' />
                                        </entity>
                                    </fetch>";

                    var query2 = new FetchExpression(fetchXml2);//fetches and makes list of account entities
                    EntityCollection results2 = service.RetrieveMultiple(query2);
                    var accountsListed = results2.Entities.ToList();

                    //makes hashtable to put account information in
                    Hashtable accountsHT = new Hashtable();
                    results2.Entities.ToList().ForEach(x => {
                        if (x.Attributes.Contains("name") && x.Attributes.Contains("accountid"))
                        {
                            string ha = x.Attributes["name"].ToString();
                            string he = x.Attributes["accountid"].ToString();
                            if (!accountsHT.Contains(ha))
                            {
                                accountsHT.Add(ha, he);
                            }

                        }
                    });//constructs hashtable to check 


                    EntityReference accRefNew;


                    //checks if name is contained in hashtable
                    bool nameCheck = check(desire, accountsHT);
                    
                    if (nameCheck)//if account exists, it simply makes the entityreference and bonds it to
                    {
                        string accGuidString =(string) accountsHT[desire];
                        accRefNew = new EntityReference("account", new Guid(accGuidString));

                    }
                    else
                    {//creates new account for new insurance company as well as a task that reminds the lab employee to check this information
                        var theNewAccount = new Entity("account");
                        theNewAccount["name"] = desire;
                        Guid newAccountGuid = service.Create(theNewAccount);
                        accRefNew = new EntityReference("account", newAccountGuid);


                        var followUp = new Entity("task");
                        followUp["subject"] = "Check new account created, found in regarding account field";
                        followUp["description"] =
                     "Follow up with the customer. Check if there are any new issues that need resolution. Please verify this is not duplicate account. " +
                     "Fill in details like address, phone number etc";
                        followUp["scheduledend"] = DateTime.Now.AddDays(2);
                        followUp["regardingobjectid"] = accRefNew;
                        followUp["cr45e_needsreminder"]=true;
                        service.Create(followUp);
                    }

 


                    Guid contactGuid = contRef.Id;
                    //Guid desiredAccGuid = accRef.Id;

                    ColumnSet attributes = new ColumnSet(new string[] { "firstname", "lastname", "parentcustomerid", "contactid" });

                    contact = service.Retrieve("contact", contactGuid, attributes);

                    contact["ss_insuranceplan"] =  accRefNew;

                    tracingService.Trace("ContactPlugin: Updating contacts account id.");
                    service.Update(contact);
                    //followup["subject"] = "Send e-mail to the new customer.";
                    // followup["description"] =
                    // "Follow up with the customer. Check if there are any new issues that need resolution.";
                    //followup["scheduledstart"] = DateTime.Now.AddDays(7);
                    //followup["scheduledend"] = DateTime.Now.AddDays(7);
                    //followup["category"] = context.PrimaryEntityName;

                    // Refer to the account in the task activity.
                    // if (context.OutputParameters.Contains("id"))
                    //{
                    //  Guid regardingobjectid = new Guid(context.OutputParameters["id"].ToString());
                    //  string regardingobjectidType = "account";

                    // followup["regardingobjectid"] =
                    //  new EntityReference(regardingobjectidType, regardingobjectid);
                    // }

                    // Create the task in Microsoft Dynamics CRM.

                    
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
        public static bool check(string name, Hashtable aht)//need to account for first letter not be capitalized
        {
            //bool x = false;
            string noSpcName = name.Replace(" ", "");
            string lowerName = name.ToLower();
            string loNoSpcName = noSpcName.ToLower();
            string upperName = name.ToUpper();
            string upNoSpcName = noSpcName.ToUpper();


            if (aht.ContainsKey(name))
            {
                return true;
            }
            else if(aht.ContainsKey(noSpcName))
            {
                return true;
            }
            else if (aht.ContainsKey(lowerName))
            {
                return true;
            }
            else if (aht.ContainsKey(loNoSpcName))
            {
                return true;
            }
            else if (aht.ContainsKey(upNoSpcName))
            {
                return true;
            }
            else
            {
                return false;
            }

            //return x;
        }
    }
}
