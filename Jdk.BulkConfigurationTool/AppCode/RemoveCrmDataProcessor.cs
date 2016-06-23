using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal class RemoveCrmDataProcessor : CrmDataProcessor
    {
        public RemoveCrmDataProcessor(IOrganizationService service, ConfigurationFile input) : base(service, input)
        {
        }

        public override void ProcessData()
        {
            var successfulRequests = 0;

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
                batch.Requests.AddRange(MapDataToRequests(MapAttributes, attributeData));
                batch.Requests.AddRange(MapDataToRequests(MapOneToMany, oneToManyData));
                batch.Requests.AddRange(MapDataToRequests(MapManyToMany, manyToManyData));
                successfulRequests += ExecuteBatch(batch);
            }

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
                entitiesBatch.Requests.AddRange(MapDataToRequests(MapEntities, entityData));
                entitiesBatch.Requests.AddRange(MapDataToRequests(MapOptionSets, optionSetData));
                successfulRequests = ExecuteBatch(entitiesBatch);
            }

            if (successfulRequests > 0)
            {
                Service.Execute(new PublishAllXmlRequest());
            }
        }

        private OrganizationRequest MapOptionSets(object[] data)
        {
            var column = InputFile.Worksheets[ConfigurationFile.WorkSheets.OptionSets].Columns.First(x => x.TargetField.Equals(ConfigurationFile.OptionSetFields.SchemaName));
            var name = (data[column.Position - 1] as string).ToLower();

            return new DeleteOptionSetRequest
            {
                Name = name
            };
        }


        private OrganizationRequest MapAttributes(object[] data)
        {
            var entityColumn = InputFile.Worksheets[ConfigurationFile.WorkSheets.Attributes].Columns.First(x => x.TargetField.Equals(ConfigurationFile.AttributeFields.EntityLogicalName));
            var entityLogicalName = (data[entityColumn.Position - 1] as string).ToLower();

            var attributeColumn = InputFile.Worksheets[ConfigurationFile.WorkSheets.Attributes].Columns.First(x => x.TargetField.Equals(ConfigurationFile.AttributeFields.SchemaName));
            var attributeLogicalName = (data[attributeColumn.Position - 1] as string).ToLower();

            return new DeleteAttributeRequest
            {
                EntityLogicalName = entityLogicalName,
                LogicalName = attributeLogicalName
            };
        }

        private List<OrganizationRequest> MapDataToRequests(Func<object[], OrganizationRequest> mapper, List<object[]> data)
        {
            var requests = new List<OrganizationRequest>();
            foreach (var row in data)
            {
                try
                {
                    var request = mapper?.Invoke(row);
                    requests.Add(request);
                }
                catch (Exception e)
                {
                    OnRaiseError(e.Message);
                }
            }
            return requests;
        }

        private OrganizationRequest MapEntities(object[] data)
        {
            var column = InputFile.Worksheets[ConfigurationFile.WorkSheets.Entities].Columns.First(x => x.TargetField.Equals(ConfigurationFile.EntityFields.SchemaName));
            var entityLogicalName = (data[column.Position - 1] as string).ToLower();

            return new DeleteEntityRequest
            {
                LogicalName = entityLogicalName
            };
        }

        private OrganizationRequest MapManyToMany(object[] data)
        {
            var column = InputFile.Worksheets[ConfigurationFile.WorkSheets.ManyToManyRelationships].Columns.First(x => x.TargetField.Equals(ConfigurationFile.ManyToManyFields.SchemaName));
            var name = data[column.Position - 1] as string;

            return new DeleteRelationshipRequest
            {
                Name = name
            };
        }

        private OrganizationRequest MapOneToMany(object[] data)
        {
            var column = InputFile.Worksheets[ConfigurationFile.WorkSheets.OneToManyRelationships].Columns.First(x => x.TargetField.Equals(ConfigurationFile.OneToManyFields.SchemaName));
            var name = data[column.Position - 1] as string;

            return new DeleteRelationshipRequest
            {
                Name = name
            };
        }

    }
}
