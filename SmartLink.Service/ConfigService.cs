﻿/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Microsoft.Azure;
using System;
using System.Configuration;

namespace SmartLink.Service
{
    public class ConfigService : IConfigService
    {
        public ConfigService()
        {
        }

        public string ClientId
        {
            get { return CloudConfigurationManager.GetSetting("ida:ClientId"); }
        }

        public string ClientSecret
        {
            get { return CloudConfigurationManager.GetSetting("ida:ClientSecret"); }
        }

        public string AzureAdInstance
        {
            get { return CloudConfigurationManager.GetSetting("ida:AADInstance"); }
        }

        public string AzureAdTenantId
        {
            get { return CloudConfigurationManager.GetSetting("ida:TenantId"); }
        }

        public string GraphResourceUrl
        {
            get { return "https://graph.microsoft.com/v1.0/"; }
        }

        public string AzureAdGraphResourceURL
        {
            get { return "https://graph.microsoft.com/"; }
        }

        public string AzureAdAuthority
        {
            get { return AzureAdInstance + AzureAdTenantId; }
        }

        public string ClaimTypeObjectIdentifier
        {
            get { return "http://schemas.microsoft.com/identity/claims/objectidentifier"; }
        }

        public string AzureWebJobsStorage
        {
            get
            {
                return ConfigurationManager.ConnectionStrings["AzureWebJobsStorage"].ConnectionString;
            }
        }

        public string AzureWebJobDashboard
        {
            get
            {
                return ConfigurationManager.ConnectionStrings["AzureWebJobsDashboard"].ConnectionString;
            }
        }

        public string SendGridMessageUserName
        {
            get
            {
                return CloudConfigurationManager.GetSetting("SendGridMessageUserName");
            }
        }

        public string SendGridMessagePassword
        {
            get
            {
                return CloudConfigurationManager.GetSetting("SendGridMessagePassword");
            }
        }

        public string SendGridMessageFromAddress
        {
            get
            {
                return CloudConfigurationManager.GetSetting("SendGridMessageFromAddress");
            }
        }

        public string SendGridMessageFromDisplayName
        {
            get
            {
                return CloudConfigurationManager.GetSetting("SendGridMessageFromDisplayName");
            }
        }

        public string[] SendGridMessageToAddress
        {
            get
            {
                return CloudConfigurationManager.GetSetting("SendGridMessageToAddress").Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            }
        }

        public string WebJobClientId
        {
            get { return CloudConfigurationManager.GetSetting("ida:WebJobClientId"); }
        }

        public string SharePointUrl
        {
            get
            {
                return CloudConfigurationManager.GetSetting("SharePointUrl");
            }
        }

        public string CertificatePassword
        {
            get
            {
                return CloudConfigurationManager.GetSetting("CertificatePassword");
            }
        }

        public string CertificateFile
        {
            get
            {

                return CloudConfigurationManager.GetSetting("CertificateFile");
            }
        }
    }
}