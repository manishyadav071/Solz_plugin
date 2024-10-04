using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

using System;

namespace Solzit.Plugins
{
	public class ProformaInvoicenumber : IPlugin
	{
		public void Execute(IServiceProvider serviceProvider)
		{
			// Obtain the tracing service
			ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
			tracingService.Trace("ProformaInvoicenumber plugin execution started.");

			// Obtain the execution context
			IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
			tracingService.Trace("Execution context obtained.");

			try
			{
				// Check if the target entity is Proforma Invoice and status is being updated or created
				if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity entity)
				{
					tracingService.Trace($"Target entity: {entity.LogicalName}");

					// Ensure the target entity is an invoice
					if (entity.LogicalName == "invoice")
					{
						IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
						IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

						// Check if it's an update
						bool isUpdate = context.MessageName == "Update";
						tracingService.Trace($"Message Name: {context.MessageName}");

						// Ensure status code is being updated to "Proforma"
						if (isUpdate && entity.Contains("statuscode"))
						{
							OptionSetValue statusCode = entity.GetAttributeValue<OptionSetValue>("statuscode");
							tracingService.Trace($"Status code: {statusCode?.Value}");

							if (statusCode != null && statusCode.Value == 674180001) // "Proforma" status value
							{
								tracingService.Trace("Status is 'Proforma'. Proceeding with auto-numbering.");

								// Check if Proforma Number already exists
								if (!entity.Contains("solz_proformainvoicenumber") || string.IsNullOrWhiteSpace(entity.GetAttributeValue<string>("solz_proformainvoicenumber")))
								{
									tracingService.Trace("Proforma Invoice number not present, generating new one.");

									// Fetch Auto Counter values
									Guid businessUnitId = Guid.Empty;
									Entity InvoiceData = service.Retrieve("invoice", entity.Id, new ColumnSet("solz_bussinessunit"));
									 businessUnitId = InvoiceData.GetAttributeValue<EntityReference>("solz_bussinessunit").Id;
									if (businessUnitId == Guid.Empty)
									{
										tracingService.Trace("Business Unit is null, aborting operation.");
										return;
									}
									tracingService.Trace("Before fetch:");

									string fetchXml = $@"
                                        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='solz_autocounter'>
                                                <attribute name='solz_entityname' />
                                                <attribute name='solz_suffixseparator' />
                                                <attribute name='solz_suffix' />
                                                <attribute name='solz_prefixseparator' />
                                                <attribute name='solz_prefix' />
                                                <attribute name='solz_padding' />
                                                <attribute name='solz_incrementby' />
                                                <attribute name='solz_currentposition' />
                                                <attribute name='solz_bussinessunit' />
                                                <attribute name='solz_attributename' />
                                                <attribute name='solz_autocounterid' />
                                                <order attribute='solz_entityname' descending='false' />
                                                <filter type='and'>
                                                    <condition attribute='solz_bussinessunit' operator='eq' value='{businessUnitId}' />
                                                    <condition attribute='solz_attributename' operator='eq' value='solz_proformainvoicenumber' />
                                                </filter>
                                            </entity>
                                        </fetch>";
									
									tracingService.Trace("Fetching auto counter data."+fetchXml);
									EntityCollection autoCounterRecords = service.RetrieveMultiple(new FetchExpression(fetchXml));
									tracingService.Trace(fetchXml);

									if (autoCounterRecords.Entities.Count > 0)
									{
										Entity autoCounter = autoCounterRecords.Entities[0];
										tracingService.Trace("Auto counter record retrieved.");

										// Retrieve auto-counter values
										int currentPosition = autoCounter.GetAttributeValue<int>("solz_currentposition");
										int incrementBy = autoCounter.GetAttributeValue<int>("solz_incrementby");
										int padding = autoCounter.GetAttributeValue<int>("solz_padding");
										string prefix = autoCounter.GetAttributeValue<string>("solz_prefix");
										string prefixSeparator = autoCounter.GetAttributeValue<string>("solz_prefixseparator");
										string suffix = autoCounter.GetAttributeValue<string>("solz_suffix");
										string suffixSeparator = autoCounter.GetAttributeValue<string>("solz_suffixseparator");

										// Increment the current position
										currentPosition += incrementBy;

										// Generate the new Proforma Invoice Number
										string newProformaNumber = (prefix ?? "") + (prefixSeparator ?? "") + currentPosition.ToString().PadLeft(padding, '0') + (suffixSeparator ?? "") + (suffix ?? "");
										tracingService.Trace($"New Proforma Number: {newProformaNumber}");

										// Update Proforma Invoice and Auto Counter
										entity["solz_proformainvoicenumber"] = newProformaNumber;
										autoCounter["solz_currentposition"] = currentPosition;

										// Use the service to update the autoCounter first
										service.Update(autoCounter);

										// Set the Proforma Invoice Date
										entity["solz_proformainvoicedate"] = DateTime.UtcNow;
										service.Update(entity);
									}
									else
									{
										tracingService.Trace("No auto counter records found.");
									}
								}
								else
								{
									tracingService.Trace("Proforma Invoice number already exists.");
								}
							}else if(statusCode != null && statusCode.Value == 4)
							{
								if (!entity.Contains("solz_solzinvoicenumber") || string.IsNullOrWhiteSpace(entity.GetAttributeValue<string>("solz_solzinvoicenumber")))
								{
									tracingService.Trace("Solz Invoice number not present, generating new one.");

									// Fetch Auto Counter values
									Guid businessUnitId = Guid.Empty;
									Entity InvoiceData = service.Retrieve("invoice", entity.Id, new ColumnSet("solz_bussinessunit"));
									businessUnitId = InvoiceData.GetAttributeValue<EntityReference>("solz_bussinessunit").Id;
									if (businessUnitId == Guid.Empty)
									{
										tracingService.Trace("Business Unit is null, aborting operation.");
										return;
									}
									tracingService.Trace("Before fetch:");

									string fetchXml = $@"
                                        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='solz_autocounter'>
                                                <attribute name='solz_entityname' />
                                                <attribute name='solz_suffixseparator' />
                                                <attribute name='solz_suffix' />
                                                <attribute name='solz_prefixseparator' />
                                                <attribute name='solz_prefix' />
                                                <attribute name='solz_padding' />
                                                <attribute name='solz_incrementby' />
                                                <attribute name='solz_currentposition' />
                                                <attribute name='solz_bussinessunit' />
                                                <attribute name='solz_attributename' />
                                                <attribute name='solz_autocounterid' />
                                                <order attribute='solz_entityname' descending='false' />
                                                <filter type='and'>
                                                    <condition attribute='solz_bussinessunit' operator='eq' value='{businessUnitId}' />
                                                    <condition attribute='solz_attributename' operator='eq' value='solz_solzinvoicenumber' />
                                                </filter>
                                            </entity>
                                        </fetch>";

									tracingService.Trace("Fetching auto counter data." + fetchXml);
									EntityCollection autoCounterRecords = service.RetrieveMultiple(new FetchExpression(fetchXml));
									tracingService.Trace(fetchXml);

									if (autoCounterRecords.Entities.Count > 0)
									{
										Entity autoCounter = autoCounterRecords.Entities[0];
										tracingService.Trace("Auto counter record retrieved.");

										// Retrieve auto-counter values
										int currentPosition = autoCounter.GetAttributeValue<int>("solz_currentposition");
										int incrementBy = autoCounter.GetAttributeValue<int>("solz_incrementby");
										int padding = autoCounter.GetAttributeValue<int>("solz_padding");
										string prefix = autoCounter.GetAttributeValue<string>("solz_prefix");
										string prefixSeparator = autoCounter.GetAttributeValue<string>("solz_prefixseparator");
										string suffix = autoCounter.GetAttributeValue<string>("solz_suffix");
										string suffixSeparator = autoCounter.GetAttributeValue<string>("solz_suffixseparator");

										// Increment the current position
										currentPosition += incrementBy;

										// Generate the new Proforma Invoice Number
										string newProformaNumber = (prefix ?? "") + (prefixSeparator ?? "") + currentPosition.ToString().PadLeft(padding, '0') + (suffixSeparator ?? "") + (suffix ?? "");
										tracingService.Trace($"New SOLZ invoice Number: {newProformaNumber}");

										// Update Proforma Invoice and Auto Counter
										entity["solz_solzinvoicenumber"] = newProformaNumber;
										autoCounter["solz_currentposition"] = currentPosition;

										// Use the service to update the autoCounter first
										service.Update(autoCounter);

										// Set the Proforma Invoice Date
										//entity["solz_proformainvoicedate"] = DateTime.UtcNow;
										service.Update(entity);
									}
									else
									{
										tracingService.Trace("No auto counter records found.");
									}
								}
								else
								{
									tracingService.Trace("solz Invoice number already exists.");
								}
							}

						}
					}
				}
			}
			catch (Exception ex)
			{
				tracingService.Trace($"Exception: {ex.Message}");
				throw new InvalidPluginExecutionException($"An error occurred in the ProformaInvoicenumber plugin: {ex.Message}", ex);
			}

			tracingService.Trace("ProformaInvoicenumber plugin execution finished.");
		}
	}
}
