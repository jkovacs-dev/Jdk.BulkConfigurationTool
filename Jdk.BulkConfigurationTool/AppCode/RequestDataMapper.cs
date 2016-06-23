using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal abstract class RequestDataMapper
    {
        private List<ConfigurationFile.Column> orderedColumns;

        protected RequestDataMapper(List<ConfigurationFile.Column> orderedColumns, int lcId)
        {
            Columns = orderedColumns.Where(x => x.Position > 0);
            LcId = lcId;
        }

        protected IEnumerable<ConfigurationFile.Column> Columns { get; private set; }
        protected int LcId { get; private set; }

        internal abstract OrganizationRequest Map(object[] dataRowValues);

        protected ConfigurationFile.Column GetColumn(Enum targetField) => Columns.First(x => targetField.Equals(x.TargetField));

        protected OptionMetadataCollection ParseOptions(string options)
        {
            var result = new OptionMetadataCollection();
            foreach (var option in options.Split(';'))
            {
                var tokens = option.Split('|');
                if(tokens.Length > 1)
                {
                    var metadata = new OptionMetadata
                    {
                        Label = new Label(tokens[0], LcId),
                        Value = Convert.ToInt32(tokens[1])
                    };
                    result.Add(metadata);
                }
            }
            return result;
        }
    }
}
