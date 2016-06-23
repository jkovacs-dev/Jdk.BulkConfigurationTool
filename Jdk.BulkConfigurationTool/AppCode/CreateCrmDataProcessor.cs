using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal class CreateCrmDataProcessor : CrmDataProcessor
    {
        public CreateCrmDataProcessor(IOrganizationService service, ConfigurationFile input) : base(service, input)
        {
            var userInfo = (WhoAmIResponse)Service.Execute(new WhoAmIRequest());
            OrgLcId = RetrieveOrgUiLanguageCode(userInfo.OrganizationId);
        }

        protected int OrgLcId { get; private set; }

        public override void ProcessData()
        {
            var successfulRequests = 0;
            var optionSetData = InputFile.Worksheets[ConfigurationFile.WorkSheets.OptionSets].Data;
            var entityData = InputFile.Worksheets[ConfigurationFile.WorkSheets.Entities].Data;
            if (entityData.Count > 0 || optionSetData.Count > 0)
            {
                var entitiesBatch = new ExecuteMultipleRequest
                {
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = true,
                        ReturnResponses = true
                    },
                    Requests = new OrganizationRequestCollection()
                };
                var entityMapper = new CreateEntityRequestDataMapper(InputFile.Worksheets[ConfigurationFile.WorkSheets.Entities].Columns, OrgLcId);
                entitiesBatch.Requests.AddRange(MapDataToRequests(entityMapper, entityData));
                var optionSetMapper = new CreateOptionSetRequestDataMapper(InputFile.Worksheets[ConfigurationFile.WorkSheets.OptionSets].Columns, OrgLcId);
                entitiesBatch.Requests.AddRange(MapDataToRequests(optionSetMapper, optionSetData));
                successfulRequests = ExecuteBatch(entitiesBatch);
            }
            var attributeData = InputFile.Worksheets[ConfigurationFile.WorkSheets.Attributes].Data;
            var oneToManyData = InputFile.Worksheets[ConfigurationFile.WorkSheets.OneToManyRelationships].Data;
            var manyToManyData = InputFile.Worksheets[ConfigurationFile.WorkSheets.ManyToManyRelationships].Data;

            if (attributeData.Count > 0 || oneToManyData.Count > 0 || manyToManyData.Count > 0)
            {
                var batch = new ExecuteMultipleRequest
                {
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = true,
                        ReturnResponses = true
                    },
                    Requests = new OrganizationRequestCollection()
                };
                var attributeMapper = new CreateAttributeRequestDataMapper(InputFile.Worksheets[ConfigurationFile.WorkSheets.Attributes].Columns, OrgLcId);
                batch.Requests.AddRange(MapDataToRequests(attributeMapper, attributeData));
                var oneToManyMapper = new CreateOneToManyRequestDataMapper(InputFile.Worksheets[ConfigurationFile.WorkSheets.OneToManyRelationships].Columns, OrgLcId);
                batch.Requests.AddRange(MapDataToRequests(oneToManyMapper, oneToManyData));
                var manyToManyMapper = new CreateManyToManyRequestDataMapper(InputFile.Worksheets[ConfigurationFile.WorkSheets.ManyToManyRelationships].Columns, OrgLcId);
                batch.Requests.AddRange(MapDataToRequests(manyToManyMapper, manyToManyData));
                successfulRequests += ExecuteBatch(batch);
            }
            if (successfulRequests > 0)
            {
                Service.Execute(new PublishAllXmlRequest());
            }
        }

        private List<OrganizationRequest> MapDataToRequests(RequestDataMapper mapper, List<object[]> data)
        {
            var requests = new List<OrganizationRequest>();
            foreach (var row in data)
            {
                try
                {
                    var request = mapper.Map(row);
                    requests.Add(request);
                }
                catch (Exception e)
                {
                    OnRaiseError(e.Message);
                }
            }
            return requests;
        }

        private int RetrieveOrgUiLanguageCode(Guid userId)
        {
            var orgLcIdQuery = new QueryExpression("organization");
            orgLcIdQuery.ColumnSet.AddColumns("languagecode");
            orgLcIdQuery.Criteria.AddCondition("organizationid", ConditionOperator.Equal, userId);
            var queryResult = Service.RetrieveMultiple(orgLcIdQuery);
            if (queryResult.Entities.Count > 0)
            {
                return (int)queryResult.Entities[0]["languagecode"];
            }
            return 0;
        }

        [ResponseLogHandler(typeof(CreateEntityResponse))]
        private static string HandleCreateEntityResponse(OrganizationRequest request)
        {
            var createRequest = (CreateEntityRequest)request;

            var result = $"Created entity {createRequest.Entity.SchemaName}";
            if(!string.IsNullOrEmpty(createRequest.SolutionUniqueName))
            {
                result += $" in solution {createRequest.SolutionUniqueName}";
            }
            result += ".";
            return result;
        }

        [ResponseLogHandler(typeof(CreateAttributeResponse))]
        private static string HandleCreateAttributeResponse(OrganizationRequest request)
        {
            var createRequest = (CreateAttributeRequest)request;

            var result = $"Added attribute {createRequest.Attribute.SchemaName} to {createRequest.EntityName}";
            if (!string.IsNullOrEmpty(createRequest.SolutionUniqueName))
            {
                result += $" in solution {createRequest.SolutionUniqueName}";
            }
            result += ".";
            return result;
        }

        [ResponseLogHandler(typeof(CreateOptionSetResponse))]
        private static string HandleCreateOptionSetResponse(OrganizationRequest request)
        {
            var createRequest = (CreateOptionSetRequest)request;

            var result = $"Created global option set {createRequest.OptionSet.Name} with {((OptionSetMetadata)createRequest.OptionSet).Options.Count} options";
            if (!string.IsNullOrEmpty(createRequest.SolutionUniqueName))
            {
                result += $" in solution {createRequest.SolutionUniqueName}";
            }
            result += ".";
            return result;
        }

        [ResponseLogHandler(typeof(CreateOneToManyResponse))]
        private static string HandleCreateOneToManyResponse(OrganizationRequest request)
        {
            var createRequest = (CreateOneToManyRequest)request;

            var result = $"Added 1:N relationship {createRequest.OneToManyRelationship.SchemaName} from {createRequest.OneToManyRelationship.ReferencingEntity} to {createRequest.OneToManyRelationship.ReferencedEntity}";
            if (!string.IsNullOrEmpty(createRequest.SolutionUniqueName))
            {
                result += $" in solution {createRequest.SolutionUniqueName}";
            }
            result += ".";
            return result;
        }

        [ResponseLogHandler(typeof(CreateManyToManyResponse))]
        private static string HandleCreateManyToManyResponse(OrganizationRequest request)
        {
            var createRequest = (CreateManyToManyRequest)request;

            var result = $"Added N:N relationship {createRequest.ManyToManyRelationship.SchemaName} between {createRequest.ManyToManyRelationship.Entity1LogicalName} and {createRequest.ManyToManyRelationship.Entity2LogicalName}";
            if (!string.IsNullOrEmpty(createRequest.SolutionUniqueName))
            {
                result += $" in solution {createRequest.SolutionUniqueName}";
            }
            result += ".";
            return result;
        }

    }
}
