using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace Event_Plugin_Hub
{
    public class Event_Create_FollowUpPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // obtain tracing service
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var sw = Stopwatch.StartNew();
            //get execution context
            IPluginExecutionContext executionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (executionContext.InputParameters.Contains("Target") && executionContext.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)executionContext.InputParameters["Target"];

                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(executionContext.UserId);


                try
                {
                    //Create a task activity to follow up with the account in 7 days from the creation of the account
                    Entity followUp = new Entity("nf_procodetask");

                    followUp["subject"] = "send an email to new customer";
                    followUp["description"] = "check if there are issue";
                    followUp["scheduledstart"] = DateTime.Now.AddDays(7);
                    followUp["scheduledend"] = DateTime.Now.AddDays(7);
                    

                    //set new task to newly created Account
                    if (executionContext.OutputParameters.Contains("id"))
                    {
                        Guid objId = new Guid(executionContext.OutputParameters["id"].ToString());
                        string regardingobjType = "nf_event";

                        followUp["regardingobjectid"] = new EntityReference(regardingobjType, objId);
                    }

                    //create the task in power apps
                    tracingService.Trace("followUpPlugin: Creating the task activity");
                    service.Create(followUp);
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("Errore in followUp plugin", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("followUpPlugin", ex.ToString());
                }
                finally
                {
                    sw.Stop();
                    tracingService.Trace($"took: {sw.ElapsedMilliseconds} ms");
                }

            }

        }
    }
}
