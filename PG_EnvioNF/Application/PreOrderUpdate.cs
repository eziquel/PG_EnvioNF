using Microsoft.Xrm.Sdk;
using PG_EnvioNF.Logic;
using System;

namespace PG_EnvioNF
{
    public class PreOrderUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(null);


            if (context.PostEntityImages.Contains("postImage") && context.PostEntityImages["postImage"] is Entity)
            {
                Entity orderImage = context.PostEntityImages["postImage"] as Entity; 

                OrderBusinessLogic businessLogic = new OrderBusinessLogic(service);
                businessLogic.EnviarNF(orderImage);
            }
        }
    }
}