using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;

namespace Jdk.BulkConfigurationTool.AppCode
{
    class CreateOptionSetRequestDataMapper : RequestDataMapper
    {
        public CreateOptionSetRequestDataMapper(List<ConfigurationFile.Column> orderedColumns, int lcId) : base(orderedColumns, lcId)
        {
        }

        internal override OrganizationRequest Map(object[] dataRowValues)
        {
            var request = new CreateOptionSetRequest
            {
                OptionSet = new OptionSetMetadata
                {
                    IsGlobal = true,
                    OptionSetType = OptionSetType.Picklist
                }
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.OptionSetFields)column.TargetField;

                if (value != null)
                {
                    switch (field)
                    {
                        case ConfigurationFile.OptionSetFields.SolutionUniqueName:
                            request.SolutionUniqueName = value as string;
                            break;
                        case ConfigurationFile.OptionSetFields.SchemaName:
                            request.OptionSet.Name = value as string;
                            break;
                        case ConfigurationFile.OptionSetFields.DisplayName:
                            request.OptionSet.DisplayName = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.OptionSetFields.Description:
                            request.OptionSet.Description = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.OptionSetFields.Options:
                            var options = ParseOptions(value as string);
                            ((OptionSetMetadata)request.OptionSet).Options.AddRange(options);
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }

            return request;
        }
    }
}
