using System;
using Microsoft.Xrm.Sdk;
using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using System.Collections.Generic;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal class CreateEntityRequestDataMapper : RequestDataMapper
    {
        internal CreateEntityRequestDataMapper(List<ConfigurationFile.Column> orderedColumns, int lcId) : base(orderedColumns, lcId)
        {
        }

        internal override OrganizationRequest Map(object[] dataRowValues)
        {
            var request = new CreateEntityRequest
            {
                Entity = new EntityMetadata
                {
                    OwnershipType = OwnershipTypes.UserOwned
                },
                PrimaryAttribute = new StringAttributeMetadata
                {
                    FormatName = StringFormatName.Text,
                    SchemaName = EnumUtils.GetDefault(ConfigurationFile.EntityFields.PrimaryAttributeSchemaName) as string,
                    DisplayName = new Label(EnumUtils.GetDefault(ConfigurationFile.EntityFields.PrimaryAttributeDisplayName) as string, 1033),
                    MaxLength = (int)EnumUtils.GetDefault(ConfigurationFile.EntityFields.PrimaryAttributeMaxLength),
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.ApplicationRequired)
                }
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.EntityFields)column.TargetField;

                if (value != null)
                {
                    switch (field)
                    {
                        case ConfigurationFile.EntityFields.SolutionUniqueName:
                            request.SolutionUniqueName = value as string;
                            break;
                        case ConfigurationFile.EntityFields.SchemaName:
                            request.Entity.SchemaName = value as string;
                            break;
                        case ConfigurationFile.EntityFields.DisplayName:
                            request.Entity.DisplayName = new Label(value as string, 1033);
                            break;
                        case ConfigurationFile.EntityFields.DisplayCollectionName:
                            request.Entity.DisplayCollectionName = new Label(value as string, 1033);
                            break;
                        case ConfigurationFile.EntityFields.Description:
                            request.Entity.Description = new Label(value as string, 1033);
                            break;
                        case ConfigurationFile.EntityFields.EntityColor:
                            request.Entity.EntityColor = value as string;
                            break;
                        case ConfigurationFile.EntityFields.OwnershipType:
                            request.Entity.OwnershipType = (OwnershipTypes)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.PrimaryAttributeSchemaName:
                            request.PrimaryAttribute.SchemaName = value as string;
                            break;
                        case ConfigurationFile.EntityFields.PrimaryAttributeDisplayName:
                            request.PrimaryAttribute.DisplayName = new Label(value as string, 1033);
                            break;
                        case ConfigurationFile.EntityFields.PrimaryAttributeMaxLength:
                            request.PrimaryAttribute.MaxLength = Convert.ToInt32(value);
                            break;
                        case ConfigurationFile.EntityFields.PrimaryAttributeDescription:
                            request.PrimaryAttribute.Description = new Label(value as string, 1033);
                            break;
                        case ConfigurationFile.EntityFields.HasActivities:
                            request.HasActivities = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.HasNotes:
                            request.HasNotes = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.IsConnectionsEnabled:
                            request.Entity.IsConnectionsEnabled = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.EntityFields.IsBusinessProcessEnabled:
                            request.Entity.IsBusinessProcessEnabled = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.IsMailMergeEnabled:
                            request.Entity.IsMailMergeEnabled = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.EntityFields.IsDocumentManagementEnabled:
                            request.Entity.IsDocumentManagementEnabled = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.IsValidForQueue:
                            request.Entity.IsValidForQueue = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.EntityFields.IsKnowledgeManagementEnabled:
                            request.Entity.IsKnowledgeManagementEnabled = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.IsQuickCreateEnabled:
                            request.Entity.IsQuickCreateEnabled = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.IsDuplicateDetectionEnabled:
                            request.Entity.IsDuplicateDetectionEnabled = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.EntityFields.IsAuditEnabled:
                            request.Entity.IsAuditEnabled = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.EntityFields.ChangeTrackingEnabled:
                            request.Entity.ChangeTrackingEnabled = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.IsVisibleInMobile:
                            request.Entity.IsVisibleInMobile = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.EntityFields.IsVisibleInMobileClient:
                            request.Entity.IsVisibleInMobileClient = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.EntityFields.IsReadOnlyInMobileClient:
                            request.Entity.IsReadOnlyInMobileClient = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.EntityFields.IsAvailableOffline:
                            request.Entity.IsAvailableOffline = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.EntityHelpUrlEnabled:
                            request.Entity.EntityHelpUrlEnabled = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.EntityFields.EntityHelpUrl:
                            request.Entity.EntityHelpUrl = value as string;
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
