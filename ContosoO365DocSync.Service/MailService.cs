/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading.Tasks;

namespace ContosoO365DocSync.Service
{
    public class MailService : IMailService
    {
        private readonly IConfigService _configService;

        private NetworkCredential _credentials;

        public MailService(IConfigService configService)
        {
            _configService = configService;
            _credentials = new NetworkCredential(_configService.SendGridMessageUserName, _configService.SendGridMessagePassword);
        }

        /// <summary>
        /// Send plain text email.
        /// </summary>
        /// <param name="fromAddress"></param>
        /// <param name="fromDisplayName"></param>
        /// <param name="toAddresses"></param>
        /// <param name="subject"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        public async Task SendPlainTextMailAsync(string fromAddress, string fromDisplayName, IEnumerable<string> toAddresses, string subject, string content)
        {
            MailMessage mailMsg = new MailMessage();

            // To
            foreach (var toAddress in toAddresses)
            {
                mailMsg.To.Add(toAddress);
            }

            // From
            mailMsg.From = new MailAddress(fromAddress, fromDisplayName);

            // Subject and multipart/alternative Body
            mailMsg.Subject = subject;
            string text = content;
            mailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(text, null, MediaTypeNames.Text.Plain));

            // Init SmtpClient and send
            SmtpClient smtpClient = new SmtpClient("smtp.sendgrid.net", Convert.ToInt32(587));
            smtpClient.Credentials = _credentials;

            await smtpClient.SendMailAsync(mailMsg);
        }
    }
}