/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Autofac;
using Microsoft.Azure.WebJobs.Host;

namespace ContosoO365DocSync.WebJob
{
    public class AutofacJobActivator : IJobActivator
    {
        private readonly IContainer _container;

        public AutofacJobActivator(IContainer container)
        {
            _container = container;
        }

        public T CreateInstance<T>()
        {
            return _container.Resolve<T>();
        }
    }
}