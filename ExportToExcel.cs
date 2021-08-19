using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Services
{
    public class ExportToExcelLog : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            var parentContext = context.ParentContext;
            var initiatingUser = context.InitiatingUserId;

            if (parentContext != null && (parentContext.MessageName == "ExportToExcel" || parentContext.MessageName == "ExportDynamicToExcel"))
            {
                if (context.InputParameters.Contains("Query") && ((QueryExpression)(context.InputParameters["Query"])).PageInfo.PageNumber > 1)
                    return;

                IDictionary<String, String> queryInfo = getExportInformation(parentContext, new Dictionary<String, String>());
                Guid systemUser = getUserId(service, ADMINUSERNAME);
                IOrganizationService _adminService = serviceFactory.CreateOrganizationService(systemUser);

                Entity excelLog = new Entity("EntityLogicalName");
                excelLog["new_name"] = "Export_" + queryInfo["entityName"] + "_" +  queryInfo["viewTitle"];
                excelLog["new_user"] = new EntityReference(CrmSystemuser.EntityLogicalName, initiatingUser);
                excelLog["new_entityname"] = queryInfo["entityName"] ;
                excelLog["new_exportedviewname"] =  queryInfo["viewTitle"];
                excelLog["new_exportedcolumns"] = queryInfo["layoutXml"];
                excelLog["new_usedfiltre"] = queryInfo["fetchXmlFilter"];
                excelLog["new_exportdate"] = DateTime.Now;

                _adminService.Create(excelLog);
            }
        }

        private Dictionary<String, String> getExportInformation(IPluginExecutionContext parentContext, Dictionary<String, String> queryInfo)
        {
            if (parentContext != null && parentContext.InputParameters.Contains("QueryParameters"))
            {
                var queryParameters = ((InputArgumentCollection)parentContext.InputParameters["QueryParameters"]);

                var entityName = queryParameters.Arguments.Contains("entitydisplayname") ? queryParameters.Arguments["entitydisplayname"].ToString() : String.Empty;
                queryInfo.Add("entityName", entityName);
                var viewTitle = queryParameters.Arguments.Contains("viewTitle") ? queryParameters.Arguments["viewTitle"].ToString() : String.Empty;
                queryInfo.Add("viewTitle", viewTitle);
                var layoutXml = queryParameters.Arguments.Contains("layoutXml") ? queryParameters.Arguments["layoutXml"].ToString() : String.Empty;
                queryInfo.Add("layoutXml", layoutXml);
                var fetchXmlFilter = queryParameters.Arguments.Contains("fetchXmlForFilters") ? queryParameters.Arguments["fetchXmlForFilters"].ToString() : String.Empty;
                queryInfo.Add("fetchXmlFilter", fetchXmlFilter);
            }

            return queryInfo;
        }

        private Guid getUserId(IOrganizationService service, string userName)
        {
            Guid systemUserId = Guid.Empty;

            QueryByAttribute queryByAttribute = new QueryByAttribute("systemuser");
            ColumnSet columns = new ColumnSet("systemUserid");
            queryByAttribute.AddAttributeValue("fullname", userName);
            EntityCollection retrieveUser = service.RetrieveMultiple(queryByAttribute);
            systemUserId = ((Entity)retrieveUser.Entities[0]).Id;

            return systemUserId;
        }
    }
}
