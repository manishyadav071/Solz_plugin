using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Linq.Expressions;


namespace Employeeleave
{
	public class EmployeeLeaveApplication : IPlugin
	{
		/// <summary>
		/// In leaves Application entity when we select the leaves start date to end date,calculate the total absent date and total leaves days.....
		/// </summary>
		/// 
		public void Execute(IServiceProvider serviceProvider)
		{
			try
			{
				// Obtain the execution context from the service provider
				ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
				IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
				IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
				IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

				// Check if the plugin is triggered by create or update message and the target entity is of the expected type
				if ((context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update") && context.PrimaryEntityName.ToLower() == "solz_leaveapplication")
				{
					// Get the target entity from the context


					Entity targetEntity = (Entity)context.InputParameters["Target"];

					DateTime startdate = targetEntity.GetAttributeValue<DateTime>("solz_leavedate");
					DateTime enddate = targetEntity.GetAttributeValue<DateTime>("solz_leaveenddate");
					tracingService.Trace("Start");
					Guid Emp = targetEntity.Contains("solz_employee") ? targetEntity.GetAttributeValue<EntityReference>("solz_employee").Id : Guid.Empty;
					tracingService.Trace("After Guid");
					string fetchXml1 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
					 <entity name='solz_leaveapplication'>
					 <attribute name='solz_leaveapplicationid'/>
					 <attribute name='solz_name'/>
					 <attribute name='createdon'/>
				     <attribute name='solz_leavetype'/>
					 <attribute name='solz_leavedate'/>
					 <attribute name='solz_isavailable'/>
					 <attribute name='solz_employee'/>
				     <attribute name='solz_approvedby'/>
					 <attribute name='solz_appliedon'/>
					 <attribute name='solz_numberofdays'/>
					 <attribute name='solz_leaveenddate'/>
					 <attribute name='solz_halfdayleavetype'/>
					 <order attribute='solz_appliedon' descending='true'/>
					 <filter type='and'>
					 <condition attribute='solz_employee' operator='eq'  value='{{0}}'/>
					 <condition attribute='statuscode' operator='eq' value='1'/>
					 </filter>
					 </entity>
					 </fetch>";
					fetchXml1 = string.Format(fetchXml1, Emp);
					tracingService.Trace("FetchXML" + fetchXml1);
					EntityCollection LeaveApplication = service.RetrieveMultiple(new FetchExpression(fetchXml1));
					if (LeaveApplication.Entities.Count > 1)
					{
						foreach (Entity entity in LeaveApplication.Entities)

						{

							DateTime enddate1 = entity.GetAttributeValue<DateTime>("solz_leaveenddate");
							string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
							 <entity name='solz_soluzioneholiday'>
							 <attribute name='solz_name' />
							 <attribute name='solz_holidaytype' />
							 <attribute name='solz_date' />
							 <attribute name='solz_soluzioneholidayid' />
							 <order attribute='solz_date' descending='false' />
							 <filter type='and'>
							 <condition attribute='solz_holidaytype' operator='eq' value='674180000' />
							 <condition attribute='solz_date' operator='on-or-after' value='{enddate1.ToString("yyyy-MM-dd")}'/>
							 <condition attribute='solz_date' operator='on-or-before' value='{enddate1.AddDays(1).ToString("yyyy-MM-dd")}'/>
							 </filter>
							 </entity>
							 </fetch>";

							tracingService.Trace("FetchXML" + fetchXml);
							EntityCollection fixedholiday = service.RetrieveMultiple(new FetchExpression(fetchXml));
							tracingService.Trace("Fixed Holidays Count: " + fixedholiday.Entities.Count);

							//count the number of fixed leaves days.
							int fixedholidaycount = fixedholiday.Entities.Count;
							if (fixedholidaycount > 0)
							{
								if (entity.GetAttributeValue<DateTime>("solz_leavedate") == startdate)
								{
									continue;
								}
								else if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").DayOfWeek == DayOfWeek.Saturday)
								{
									if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").AddDays(3) >= startdate)
									{
										throw new InvalidPluginExecutionException("Please Modified Previous Leave day");
									}
								}
								else if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").DayOfWeek == DayOfWeek.Sunday){
									if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").AddDays(2) >= startdate){
										throw new InvalidPluginExecutionException("Please Modified Previous Leave day");
									}
								}
								else if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").DayOfWeek == DayOfWeek.Friday || entity.GetAttributeValue<DateTime>("solz_leaveenddate").DayOfWeek == DayOfWeek.Thursday)
								{
									if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").AddDays(4) >= startdate){
										throw new InvalidPluginExecutionException("Please Modified Previous Leave day");
									}
								}
								else{
									if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").AddDays(2) >= startdate){
										throw new InvalidPluginExecutionException("Please Modified Previous Leave day");

									}

								}
							}else{
								if (entity.GetAttributeValue<DateTime>("solz_leavedate") == startdate)
								{
									continue;
								}
								else if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").DayOfWeek == DayOfWeek.Saturday)
								{
									if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").AddDays(2) >= startdate)
									{
										throw new InvalidPluginExecutionException("Please Modified Previous Leave day");
									}
								}
								else if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").DayOfWeek == DayOfWeek.Sunday)
								{
									if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").AddDays(1) >= startdate)
									{
										throw new InvalidPluginExecutionException("Please Modified Previous Leave day");
									}
								}
								else if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").DayOfWeek == DayOfWeek.Friday || entity.GetAttributeValue<DateTime>("solz_leaveenddate").DayOfWeek == DayOfWeek.Thursday)
								{
									if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").AddDays(3) >= startdate)
									{
										throw new InvalidPluginExecutionException("Please Modified Previous Leave day");
									}
								}
								else
								{
									if (entity.GetAttributeValue<DateTime>("solz_leaveenddate").AddDays(1) >= startdate)
									{
										throw new InvalidPluginExecutionException("Please Modified Previous Leave day");

									}

								}
							}
						}
					}
						startdate = startdate.AddDays(0);
						enddate = enddate.AddDays(0);

						decimal total_absentdays = 0;
						int totaldays = 0;
						if (startdate <= enddate)
						{
							tracingService.Trace("Startdate is:" + startdate);
							tracingService.Trace("Enddate is:" + enddate);
							int totaltotalweekendDays = 0;

							//calculate the weekend days..
							for (var date = startdate; date <= enddate; date = date.AddDays(1))
							{
								if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
									totaltotalweekendDays++;

								tracingService.Trace("Total Days in loop" + totaltotalweekendDays);
							}
							for (var date = startdate; date <= enddate; date = date.AddDays(1.00))
							{
								total_absentdays++;
								tracingService.Trace("Total Absent Days: " + total_absentdays);
							}
							tracingService.Trace("Total Days Count" + totaltotalweekendDays);

							//Get the Fixed Holiday count
							string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
							 <entity name='solz_soluzioneholiday'>
							 <attribute name='solz_name' />
							 <attribute name='solz_holidaytype' />
							 <attribute name='solz_date' />
							 <attribute name='solz_soluzioneholidayid' />
							 <order attribute='solz_date' descending='false' />
							 <filter type='and'>
							 <condition attribute='solz_holidaytype' operator='eq' value='674180000' />
							 <condition attribute='solz_date' operator='on-or-after' value='{startdate.ToString("yyyy-MM-dd")}'/>
							 <condition attribute='solz_date' operator='on-or-before' value='{enddate.ToString("yyyy-MM-dd")}'/>
							 </filter>
							 </entity>
							 </fetch>";

							tracingService.Trace("FetchXML" + fetchXml);
							EntityCollection fixedholiday = service.RetrieveMultiple(new FetchExpression(fetchXml));
							tracingService.Trace("Fixed Holidays Count: " + fixedholiday.Entities.Count);

							//count the number of fixed leaves days.
							int fixedholidaycount = fixedholiday.Entities.Count;

							totaldays = totaltotalweekendDays - fixedholidaycount;

							// Update the total days absent field....
							targetEntity["solz_totalabsentdays"] = totaldays;

							//total number of absent days...
							targetEntity["solz_numberofdays"] = Convert.ToDecimal(total_absentdays);

							service.Update(targetEntity);
							tracingService.Trace("Total Days after deducting fixed holidays: " + totaldays);
						}
				
				}
				
			}
			catch (Exception ex)
			{
				throw new InvalidPluginExecutionException(ex.Message);
			}

        }
	}
}
