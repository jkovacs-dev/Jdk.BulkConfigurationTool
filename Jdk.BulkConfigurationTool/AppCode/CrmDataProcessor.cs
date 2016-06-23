using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Reflection;
using System.ServiceModel;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal abstract class CrmDataProcessor : IDataProcessor
    {

        internal CrmDataProcessor(IOrganizationService service, ConfigurationFile input)
        {
            InputFile = input;
            Service = service;

        }

        public event EventHandler<string> RaiseError;
        public event EventHandler<string> RaiseSuccess;
        public ConfigurationFile InputFile { get; set; }
        public IOrganizationService Service { get; set; }

        public abstract void ProcessData();

        protected int ExecuteBatch(ExecuteMultipleRequest batchRequest)
        {
            var successfulRequests = 0;
            try
            {
                var results = (ExecuteMultipleResponse)Service.Execute(batchRequest);
                foreach (var responseItem in results.Responses)
                {
                    // A valid response
                    if (responseItem.Response != null)
                    {
                        successfulRequests++;
                        var message = InvokeResponseTypeHandler(responseItem.Response, batchRequest.Requests[responseItem.RequestIndex]);
                        OnRaiseSuccess(message);
                    }
                    // an error occurred.
                    else if (responseItem.Fault != null)
                    {
                        OnRaiseError(responseItem.Fault.Message);
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> fault)
            {
                // Check if the maximum batch size has been exceeded. The maximum batch size is only included in the fault if it
                // the input request collection count exceeds the maximum batch size.
                if (fault.Detail.ErrorDetails.Contains("MaxBatchSize"))
                {
                    var maxBatchSize = Convert.ToInt32(fault.Detail.ErrorDetails["MaxBatchSize"]);
                    if (maxBatchSize < batchRequest.Requests.Count)
                    {
                        var requests = batchRequest.Requests.ToList();
                        while (requests.Count() > 0)
                        {
                            var batchSubset = new ExecuteMultipleRequest
                            {
                                Settings = new ExecuteMultipleSettings
                                {
                                    ContinueOnError = true,
                                    ReturnResponses = true
                                },
                                Requests = new OrganizationRequestCollection()
                            };
                            batchSubset.Requests.AddRange(requests.Take(maxBatchSize));
                            requests.RemoveRange(0, maxBatchSize);
                            successfulRequests += ExecuteBatch(batchSubset);
                        }
                    }
                }
            }
            return successfulRequests;
        }

        private string InvokeResponseTypeHandler(OrganizationResponse response, OrganizationRequest request)
        {
            var handler =
                (from method in GetType().GetMethods((BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.InvokeMethod))
                 let attributes = method.GetCustomAttributes(typeof(ResponseLogHandlerAttribute), true).Where(x => 
                 {
                     return ((ResponseLogHandlerAttribute)x).DataType.Equals(response.GetType());
                     })
                 where attributes != null && attributes.Count() > 0
                 select new
                 {
                     Method = method,
                     Attributes = attributes.Cast<ResponseLogHandlerAttribute>()
                 })
                .FirstOrDefault();

            if(handler == null)
            {
                return response.ToString();
            }
            return handler.Method.Invoke(null, new object[] { request }) as string;
        }

        protected virtual void OnRaiseError(string e)
        {
            RaiseError?.Invoke(this, e);
        }

        protected virtual void OnRaiseSuccess(string e)
        {
            RaiseSuccess?.Invoke(this, e);
        }

        [AttributeUsage(AttributeTargets.Method)]
        protected sealed class ResponseLogHandlerAttribute : Attribute
        {

            public ResponseLogHandlerAttribute(Type dataType)
            {
                DataType = dataType;
            }
            public Type DataType { get; private set; }
        }
    }
}
