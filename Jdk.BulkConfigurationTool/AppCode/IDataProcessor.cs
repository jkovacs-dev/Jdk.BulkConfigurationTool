using Jdk.BulkConfigurationTool.Helpers;
using System;

namespace Jdk.BulkConfigurationTool.AppCode
{
    public interface IDataProcessor
    {
        ConfigurationFile InputFile { get; set; }

        event EventHandler<string> RaiseError;
        event EventHandler<string> RaiseSuccess;

        void ProcessData();
    }
}
