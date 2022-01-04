using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace ExtendingRLS
{
    public class RetriveAnyEntitiesPreOperation : IPlugin
    {
        public const string SavedQueryLogicalName = "savedquery";

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService = (ITracingService)
                serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider 
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the service factory
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)
                        serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            // Obtain the organization service factory
            IOrganizationService organizationService = serviceFactory.CreateOrganizationService(context.UserId);

            // Avoid executing plugin for Saved Query table for preventing infinite loop.
            if (context.PrimaryEntityName == SavedQueryLogicalName)
                return;

            // The InputParameters collection contains all the data passed in the message request. 
            // If it is retrive multiple request 
            if (context.InputParameters.Contains("Query"))
            {
                try
                {
                    // Get query from input parameters and try to cast to query type
                    var thisQuery = context.InputParameters["Query"];
                    var fetchExpressionQuery = thisQuery as FetchExpression;
                    var queryExpressionQuery = thisQuery as QueryExpression;

                    if (fetchExpressionQuery != null)
                    {
                        // Parsing query
                        XDocument fetchXmlDoc = XDocument.Parse(fetchExpressionQuery.Query);

                        // Get the entity element of this query
                        var entityElement = fetchXmlDoc.Descendants("entity").FirstOrDefault();
                        // Get entity name of this query
                        var entityName = entityElement.Attributes("name").FirstOrDefault().Value;
                        // Using custom method to get all hidden views 
                        var savedViews = RetriveHiddenViews(organizationService, entityName);

                        // Add all filter from each view into current query
                        foreach (var view in savedViews)
                        {
                            // Parse query
                            var xml = XDocument.Parse(view.Attributes["fetchxml"].ToString());
                            // Get entity 
                            var fetchEntity = xml.Descendants("entity");

                            // Get filters 
                            var filters = fetchEntity.Elements("filter");
                            // Get linked entities 
                            var linkedEntities = fetchEntity.Elements("link-entity");

                            // Adding filters
                            foreach (var filter in filters)
                            {
                                entityElement.Add(filter);
                            }

                            // Adding linked entities
                            foreach (var link in linkedEntities)
                            {
                                entityElement.Add(link);
                            }
                        }

                        // Updating current query from the modified query
                        fetchExpressionQuery.Query = fetchXmlDoc.ToString();
                    }

                    if (queryExpressionQuery != null)
                    {

                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Extending RLS plugin [retrive multiple]: {0}", ex.ToString());
                    throw;
                }
            }

            // If it is retrive single entity 
            else if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
            {
                try
                {
                    // Obtain the target entity from the input parameters.
                    EntityReference entity = (EntityReference)context.InputParameters["Target"];
                    var entityName = entity.LogicalName;

                    // Create query (e.g. "inner query" for retrive multiple entities with filter by current entity id. All hidden filters will be applied. 
                    var fetchXml = $@"
                        <fetch>
                          <entity name='{entityName}'>
                            <filter>
                              <condition attribute='{entityName}id' operator='eq' value='{entity.Id.ToString().Replace("{", "").Replace("}", "")}'/>
                            </filter>
                          </entity>
                        </fetch>";

                    var query = new FetchExpression()
                    {
                        Query = fetchXml
                    };

                    var entityCollection = organizationService.RetrieveMultiple(query);
                    DataCollection<Entity> entities = entityCollection.Entities;

                    // If not found entities by this query then not all conditions applied
                    if (entities.Count == 0)
                    {
                        throw new InvalidPluginExecutionException(
                            OperationStatus.Canceled,
                            "Access to the resource is forbidden. Contact your Microsoft Power Apps administrator for assistance."
                        );
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Extending RLS plugin [retrive single]: {0}", ex.ToString());
                    throw;
                }
            }
        }

        private static DataCollection<Entity> RetriveHiddenViews(IOrganizationService organizationService, string entityName)
        {
            // Create query to SavedQuery table to get saved views for this entity 
            var query = new QueryExpression()
            {
                // Get only name and fetchxml fields
                ColumnSet = new ColumnSet("name", "fetchxml"),
                EntityName = SavedQueryLogicalName,
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression()
                        {
                            // Filter by current entity 
                            AttributeName = "returnedtypecode",
                            Operator = ConditionOperator.Equal,
                            Values = { entityName }
                        },
                        new ConditionExpression()
                        {
                            // Filter by name consideration 
                            // (e.g. /type=hidden, for example, view name as "Filter by high contract value /type=hidden").
                            AttributeName = "name",
                            Operator = ConditionOperator.Like,
                            Values = { "%/type=hidden%" }
                        },
                        new ConditionExpression()
                        {
                            // Filter by Public View view type 
                            AttributeName = "querytype",
                            Operator = ConditionOperator.Equal,
                            Values = { 0 }
                        },
                        new ConditionExpression()
                        {
                            // Filter by state. We can disable our hidden views by out-of-the-box options. 
                            AttributeName = "statecode",
                            Operator = ConditionOperator.Equal,
                            Values = { 0 }
                        }
                    }
                }
            };

            // Get views
            var retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
            var retrieveSavedQueriesResponse = (RetrieveMultipleResponse)organizationService.Execute(retrieveSavedQueriesRequest);
            DataCollection<Entity> savedViews = retrieveSavedQueriesResponse.EntityCollection.Entities;

            return savedViews;
        }
    }
}
