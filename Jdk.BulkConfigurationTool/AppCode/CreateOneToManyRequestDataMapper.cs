using System;
using System.Collections.Generic;
using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal class CreateOneToManyRequestDataMapper : RequestDataMapper
    {
        internal CreateOneToManyRequestDataMapper(List<ConfigurationFile.Column> orderedColumns, int lcId) : base(orderedColumns, lcId)
        {
        }

        internal override OrganizationRequest Map(object[] dataRowValues)
        {
            var request = new CreateOneToManyRequest
            {
                Lookup = new LookupAttributeMetadata {
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.ApplicationRequired)
                },
                OneToManyRelationship = new OneToManyRelationshipMetadata
                {
                    AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                    {
                        Behavior = AssociatedMenuBehavior.UseCollectionName,
                        Group = AssociatedMenuGroup.Details,
                        Order = 10000
                    },
                    CascadeConfiguration = new CascadeConfiguration
                    {
                        Assign = CascadeType.NoCascade,
                        Delete = CascadeType.RemoveLink,
                        Merge = CascadeType.NoCascade,
                        Reparent = CascadeType.NoCascade,
                        Share = CascadeType.NoCascade,
                        Unshare = CascadeType.NoCascade
                    }
                }
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.OneToManyFields)column.TargetField;

                if (value != null)
                {
                    switch (field)
                    {
                        case ConfigurationFile.OneToManyFields.SolutionUniqueName:
                            request.SolutionUniqueName = value as string;
                            break;
                        case ConfigurationFile.OneToManyFields.EntityLogicalName:
                            request.OneToManyRelationship.ReferencingEntity = (value as string).ToLower();
                            break;
                        case ConfigurationFile.OneToManyFields.RelatedEntityLogicalName:
                            request.OneToManyRelationship.ReferencedEntity = (value as string).ToLower();
                            break;
                        case ConfigurationFile.OneToManyFields.SchemaName:
                            request.OneToManyRelationship.SchemaName = value as string;
                            break;
                        case ConfigurationFile.OneToManyFields.LookupAttributeDisplayName:
                            request.Lookup.DisplayName = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.OneToManyFields.LookupAttributeSchemaName:
                            request.Lookup.SchemaName = value as string;
                            break;
                        case ConfigurationFile.OneToManyFields.RequiredLevel:
                            request.Lookup.RequiredLevel = new AttributeRequiredLevelManagedProperty((AttributeRequiredLevel)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.OneToManyFields.LookupAttributeDescription:
                            request.Lookup.Description = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.OneToManyFields.MenuBehavior:
                            request.OneToManyRelationship.AssociatedMenuConfiguration.Behavior = (AssociatedMenuBehavior)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.OneToManyFields.MenuGroup:
                            request.OneToManyRelationship.AssociatedMenuConfiguration.Group = (AssociatedMenuGroup)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.OneToManyFields.MenuCustomLabel:
                            request.OneToManyRelationship.AssociatedMenuConfiguration.Label = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.OneToManyFields.MenuOrder:
                            request.OneToManyRelationship.AssociatedMenuConfiguration.Order = Convert.ToInt32(value);
                            break;
                        case ConfigurationFile.OneToManyFields.CascadeAssign:
                            request.OneToManyRelationship.CascadeConfiguration.Assign = (CascadeType)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.OneToManyFields.CascadeShare:
                            request.OneToManyRelationship.CascadeConfiguration.Share = (CascadeType)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.OneToManyFields.CascadeUnshare:
                            request.OneToManyRelationship.CascadeConfiguration.Unshare = (CascadeType)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.OneToManyFields.CascadeReparent:
                            request.OneToManyRelationship.CascadeConfiguration.Reparent = (CascadeType)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.OneToManyFields.CascadeDelete:
                            request.OneToManyRelationship.CascadeConfiguration.Delete = (CascadeType)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.OneToManyFields.CascadeMerge:
                            request.OneToManyRelationship.CascadeConfiguration.Merge = (CascadeType)EnumUtils.GetSelectedOption(field, value);
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
