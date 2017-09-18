/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace ContosoO365DocSync.Service
{
    public interface IConfigService
    {
        string ClientId { get;}

        string ClientSecret { get; }

        string WebJobClientId { get; }

        string AzureADInstance { get; }

        string AzureADTenantId { get; }

        string GraphResourceUrl{ get; }

        string AzureADGraphResourceURL { get; }

        string AzureADAuthority { get; }

        string ClaimTypeObjectIdentifier { get; }

        string AzureWebJobsStorage { get; }

        string AzureWebJobDashboard { get; }

        string SharePointUrl { get; }

        string CertificateFile { get; }

        string CertificatePassword { get; }

        string SendGridMessageUserName { get; }

        string SendGridMessagePassword { get; }

        string SendGridMessageFromAddress { get; }

        string SendGridMessageFromDisplayName { get; }

        string[] SendGridMessageToAddress { get; }
   }
}