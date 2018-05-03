using System;
using System.ServiceModel;
using System.ServiceModel.Activation;

namespace Microsoft.Operations.Webservices
{
    /// <summary>
    /// Making a reference to this Factory allows your webservice to run on IIS even if other
    /// addresses are specified. This is a scenario which may only be encountered when you deploy to
    /// production, whereby other endpoints are present because of multiple sites or a shared
    /// environment. To use this in your webservice code, add the factory reference to your .SVC
    /// markup; for example: CodeBehind="v20100217.svc.cs"
    /// Factory="Microsoft.Operations.Webservices.MultipleHostsFactory" The way this works is that it
    /// effectively neutralizes (overwrites) any other addresses which may be present. for more
    /// details, refer: http://blog.ranamauro.com/2008/07/hosting-wcf-service-on-iis-site-with_25.html
    /// </summary>
    /// <remarks>
    /// NOTE: This architecture is current when the project was using .NET 3.5 but has not yet been
    ///       reviewed for .NET 4.0
    /// NOTE: DO not use for REST endpoints.
    /// History: I forget why we actually needed to use this. These days we don't use the 'multiple
    ///          address' model anymore - ... i.e. it's mainly for situations where you have the same
    /// code running under different domains. We don't (currently) use this is the newer projects
    /// like CommerceDatabase webservices where we have all the WCF endpoints with really
    /// fine-grained control using the global.asax replacement.
    /// </remarks>
    internal class MultipleHostsFactory : ServiceHostFactory
    {
        protected override ServiceHost CreateServiceHost(Type serviceType, Uri[] baseAddresses)
        {
            return base.CreateServiceHost(serviceType, new Uri[] { });
        }
    }
}