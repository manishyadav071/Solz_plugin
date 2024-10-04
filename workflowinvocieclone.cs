using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace SolzIT_ESS_Workflow
{
    public class CloneInvoiceRecord : CodeActivity
    {
        /// <summary>
        /// Executes the workflow activity, create the dublicate of the record
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered CloneInvoiceRecord.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("CloneInvoiceRecord.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // TODO: Implement your custom Workflow business logic.
                tracingService.Trace("Inside CloneInvoiceRecord Custom Workflow");
                if (context.Depth > 1)
                    throw new InvalidPluginExecutionException("Workflow called again");

                //Get all the fields of the Invoices on which the workflow is fired 
                Entity invoiceOld = service.Retrieve("invoice", context.PrimaryEntityId, new ColumnSet(true));

                //Create a new Invoice enity
                Entity invoiceNew = new Entity("invoice");

                //Variable Initialization
                Guid newInvoiceGuid = Guid.Empty;
                Guid newinvoiceProductGuid = Guid.Empty;
                var keyString = string.Empty;

                //Loop to get all the fields details on which the workflow is fired 
                foreach (string key in invoiceOld.Attributes.Keys)
                {
                    //Don't want to set these
                    if (key != "statuscode" && key != "statecode" && key != "solz_solzinvoicenumber" && key != "invoicenumber" && key != "name" && key != "datedelivered" && key != "duedate")
                    {
                        keyString += " || key: " + key;

                        switch (invoiceOld.Attributes[key].GetType().ToString())
                        {
                            case "Microsoft.Xrm.Sdk.EntityReference":
                                invoiceNew.Attributes[key] = invoiceOld.GetAttributeValue<EntityReference>(key);
                                keyString += " = " + invoiceOld.Attributes[key].GetType().ToString() + " : " + invoiceOld.GetAttributeValue<EntityReference>(key) +"\n";
                                break;
                            case "Microsoft.Xrm.Sdk.OptionSetValue":
                                invoiceNew.Attributes[key] = invoiceOld.GetAttributeValue<OptionSetValue>(key);
                                keyString += " = " + invoiceOld.Attributes[key].GetType().ToString() + " : " + invoiceOld.GetAttributeValue<OptionSetValue>(key) + "\n";
                                break;
                            case "Microsoft.Xrm.Sdk.Money":
                                invoiceNew.Attributes[key] = invoiceOld.GetAttributeValue<Money>(key).Value;
                                keyString += " = " + invoiceOld.Attributes[key].GetType().ToString() + " : " + invoiceOld.GetAttributeValue<Money>(key) + "\n";
                                break;
                            case "System.Guid":
                                //Don't set this, only attribute of type System.Guid is the Id value
                                break;
                            default:
                                invoiceNew.Attributes[key] = invoiceOld.Attributes[key];
                                keyString += " = " + invoiceOld.Attributes[key].GetType().ToString() + " : " + invoiceOld.Attributes[key] + "\n";
                                break;
                        }
                    }
                 

                        //If Invoice name then edit it by putting Clone in the begining
                    else if (key == "name")
                        invoiceNew["name"] = "(Clone) " + invoiceOld.Attributes[key];

                        //Set the Invoice date to today date
                    //else if (key == "datedelivered")
                    //    invoiceNew["datedelivered"] = DateTime.Today;

                    ////Set the Due date according to the payment terms code
                    //if (key == "paymenttermscode")
                    //{
                    //    int totalday = ((OptionSetValue)invoiceOld["paymenttermscode"]).Value;
                    //    if (totalday == 1)
                    //        invoiceNew["duedate"] = DateTime.Today.AddDays(5);
                    //    else if(totalday == 2)
                    //        invoiceNew["duedate"] = DateTime.Today.AddDays(15);
                    //    else if(totalday == 3)
                    //        invoiceNew["duedate"] = DateTime.Today.AddDays(10);
                    //    else if(totalday == 4)
                    //        invoiceNew["duedate"] = DateTime.Today.AddDays(30);
                    //    else if(totalday == 5)
                    //        invoiceNew["duedate"] = DateTime.Today;
                    //}
                    tracingService.Trace(keyString + "\n\n\n");
                    invoiceNew["solz_iseinvoicecompleted"] = false;
                    invoiceNew["solz_irnnumber"] = null;
                    invoiceNew["solz_acknumber"] = null;
                    invoiceNew["solz_ackdate"] = null;
					invoiceNew["solz_proformainvoicedate"] = null;
					invoiceNew["solz_proformainvoiceduedate"] = null;
					invoiceNew["solz_proformainvoicenumber"] = null;
				}

                tracingService.Trace("before creating invoice" + "\n");

                //Create New Invoice Successfully
                newInvoiceGuid = service.Create(invoiceNew);
                tracingService.Trace("after creating invoice" + "\n");
                tracingService.Trace("newInvoiceGuid=" + newInvoiceGuid + "\n");

                //On The basis of old invoice id get the invoice product lists
                string invoiceProductQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >
                                                    <entity name='invoicedetail' >   
                                                        <attribute name='invoicedetailid' />
                                                        <filter type='and' >
                                                            <condition attribute='invoiceid' operator='eq' value='{0}' />
                                                        </filter>
                                                    </entity>
                                                </fetch>";
                invoiceProductQuery = string.Format(invoiceProductQuery, context.PrimaryEntityId);

                EntityCollection invoiceProductEntityColl = service.RetrieveMultiple(new FetchExpression(invoiceProductQuery));

                //To Create Invoice Products for new invoice
                Entity invoiceProductNew = new Entity("invoicedetail");
                bool isPriceOverRidden = false;
                bool isProductOverRidden = false;

                tracingService.Trace("Total Invoice Products: " + invoiceProductEntityColl.Entities.Count + "\n");

                if (invoiceProductEntityColl != null && invoiceProductEntityColl.Entities.Count > 0)
                {
                    foreach (var invoiceProduct in invoiceProductEntityColl.Entities)
                    {

                        tracingService.Trace("invoiceProduct id=" + invoiceProduct.Id + "\n");
                        //Get all the fields of records on which the workflow runs
                        Entity invoiceProdOld = service.Retrieve("invoicedetail", invoiceProduct.Id, new ColumnSet(true));

                        //Get Inoice Product- Over ridden price and over ridden product
                        isPriceOverRidden = bool.Parse(invoiceProdOld["ispriceoverridden"].ToString());
                        isProductOverRidden = bool.Parse(invoiceProdOld["isproductoverridden"].ToString());

                        tracingService.Trace("isPriceOverRidden=" + isPriceOverRidden + "\n");
                        tracingService.Trace("isProductOverRidden=" + isProductOverRidden + "\n");

                        //Loop to get all the fields of the invoice product
                        foreach (string key in invoiceProdOld.Attributes.Keys)
                        {
                            string dontUseProductIdKey = string.Empty;
                            string dontUseProductDescription = string.Empty;
                            
                            //If Product is write-in-product donot set the productId, productDescription, unitOfMeasurement
                            if (isProductOverRidden)
                            {
                                dontUseProductIdKey = "productid";
                                dontUseProductDescription = "productdescription";
                                invoiceProductNew["productid"] = null;
                                invoiceProductNew["uomid"] = null;

                                //Set ProductDescription = Clone of Productdescription
                                if (invoiceProdOld.Contains("productdescription"))
                                    invoiceProductNew["productdescription"] =  invoiceProdOld["productdescription"].ToString();

                                tracingService.Trace("\nProduct id  removed fot the overridden price\n");
                            }

                            //Don't want to set these
                            if (key != "statuscode" && key != "statecode" && key != "invoiceid" && key != dontUseProductDescription && key != dontUseProductIdKey)
                            {
                                keyString += " || key: " + key;

                                switch (invoiceProdOld.Attributes[key].GetType().ToString())
                                {
                                    case "Microsoft.Xrm.Sdk.EntityReference":
                                        invoiceProductNew.Attributes[key] = invoiceProdOld.GetAttributeValue<EntityReference>(key);
                                        keyString += " = " + invoiceProdOld.Attributes[key].GetType().ToString() + " : " + invoiceProdOld.GetAttributeValue<EntityReference>(key) + "\n";
                                        break;
                                    case "Microsoft.Xrm.Sdk.OptionSetValue":
                                        invoiceProductNew.Attributes[key] = invoiceProdOld.GetAttributeValue<OptionSetValue>(key);
                                        keyString += " = " + invoiceProdOld.Attributes[key].GetType().ToString() + " : " + invoiceProdOld.GetAttributeValue<OptionSetValue>(key) + "\n";
                                        break;
                                    case "Microsoft.Xrm.Sdk.Money":
                                        invoiceProductNew.Attributes[key] = invoiceProdOld.GetAttributeValue<Money>(key).Value;
                                        keyString += " = " + invoiceProdOld.Attributes[key].GetType().ToString() + " : " + invoiceProdOld.GetAttributeValue<Money>(key) + "\n";
                                        break;
                                    case "System.Guid":
                                        //Don't set this, only attribute of type System.Guid is the Id value
                                        break;
                                    default:
                                        invoiceProductNew.Attributes[key] = invoiceProdOld.Attributes[key];
                                        keyString += " = " + invoiceProdOld.Attributes[key].GetType().ToString() + " : " + invoiceProdOld.Attributes[key] + "\n";
                                        break;
                                }
                            }
                                //Set the Invoice Id = new Invoice id
                            else if (key == "invoiceid")                            
                                invoiceProductNew["invoiceid"] = new EntityReference("invoice", newInvoiceGuid);
                            
                            tracingService.Trace(keyString + "\n\n\n");
                            tracingService.Trace(newInvoiceGuid + "New invoice id");
                        }

                        tracingService.Trace("before creating invoiceProduct" + "\n");

                        //Create Invoice product successfully
                        newinvoiceProductGuid = service.Create(invoiceProductNew);
                        tracingService.Trace("after creating invoiceProduct" + "\n");
                        tracingService.Trace("newInspectionGuid=" + newinvoiceProductGuid + "\n");

                        //If Price is overridden update the invoice product fields
                        if (isPriceOverRidden)
                        {
                            //For Udate Invoice Product
                            Entity updateInvoiceProduct = new Entity("invoicedetail");
                            tracingService.Trace("Before updating invoice product");

                            //Update the Price per unit, manual discount, tax
                            updateInvoiceProduct.Id = newinvoiceProductGuid;

                            if (invoiceProdOld.Contains("priceperunit"))
                                updateInvoiceProduct["priceperunit"] = invoiceProdOld["priceperunit"];

                            if (invoiceProdOld.Contains("manualdiscountamount"))
                                updateInvoiceProduct["manualdiscountamount"] = invoiceProdOld["manualdiscountamount"];

                            if (invoiceProdOld.Contains("tax"))
                                updateInvoiceProduct["tax"] = invoiceProdOld["tax"];

                            //Update the invoice product
                            service.Update(updateInvoiceProduct);
                            tracingService.Trace("After updating invoice product");
                        }
                    }
                }
                //throw new InvalidPluginExecutionException("Break");
            }
            catch (Exception e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw new InvalidPluginExecutionException(e.Message);
            }

            tracingService.Trace("Exiting CloneInvoiceRecord.Execute(), Correlation Id: {0}", context.CorrelationId);
        }
    }
}
