using System;
using System.Collections.Generic;
using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal class CreateManyToManyRequestDataMapper : RequestDataMapper
    {
        public CreateManyToManyRequestDataMapper(List<ConfigurationFile.Column> orderedColumns, int lcId) : base(orderedColumns, lcId)
        {
        }

        internal override OrganizationRequest Map(object[] dataRowValues)
        {
            var request = new CreateManyToManyRequest
            {
                ManyToManyRelationship = new ManyToManyRelationshipMetadata
                {
                    Entity1AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                    {
                        Behavior = AssociatedMenuBehavior.DoNotDisplay,
                        Group = AssociatedMenuGroup.Details,
                        Order = 10000
                    },
                    Entity2AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                    {
                        Behavior = AssociatedMenuBehavior.DoNotDisplay,
                        Group = AssociatedMenuGroup.Details,
                        Order = 10000
                    }
                }
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.ManyToManyFields)column.TargetField;

                if (value != null)
                {
                    switch (field)
                    {
                        case ConfigurationFile.ManyToManyFields.SolutionUniqueName:
                            request.SolutionUniqueName = value as string;
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity1LogicalName:
                            request.ManyToManyRelationship.Entity1LogicalName = (value as string).ToLower();
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity1MenuBehavior:
                            request.ManyToManyRelationship.Entity1AssociatedMenuConfiguration.Behavior = (AssociatedMenuBehavior)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity1MenuGroup:
                            request.ManyToManyRelationship.Entity1AssociatedMenuConfiguration.Group = (AssociatedMenuGroup)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity1MenuCustomLabel:
                            request.ManyToManyRelationship.Entity1AssociatedMenuConfiguration.Label = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity1MenuOrder:
                            request.ManyToManyRelationship.Entity1AssociatedMenuConfiguration.Order = Convert.ToInt32(value);
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity2LogicalName:
                            request.ManyToManyRelationship.Entity2LogicalName = (value as string).ToLower();
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity2MenuBehavior:
                            request.ManyToManyRelationship.Entity2AssociatedMenuConfiguration.Behavior = (AssociatedMenuBehavior)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity2MenuGroup:
                            request.ManyToManyRelationship.Entity2AssociatedMenuConfiguration.Group = (AssociatedMenuGroup)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity2MenuCustomLabel:
                            request.ManyToManyRelationship.Entity2AssociatedMenuConfiguration.Label = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.ManyToManyFields.Entity2MenuOrder:
                            request.ManyToManyRelationship.Entity2AssociatedMenuConfiguration.Order = Convert.ToInt32(value);
                            break;
                        case ConfigurationFile.ManyToManyFields.SchemaName:
                            request.ManyToManyRelationship.SchemaName = value as string;
                            break;
                        case ConfigurationFile.ManyToManyFields.IntersectEntitySchemaName:
                            request.IntersectEntitySchemaName = value as string;
                            break;
                        case ConfigurationFile.ManyToManyFields.Searchable:
                            request.ManyToManyRelationship.IsValidForAdvancedFind = (bool)EnumUtils.GetSelectedOption(field, value);
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

