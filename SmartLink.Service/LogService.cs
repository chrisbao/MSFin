﻿using Microsoft.ApplicationInsights;
using SmartLink.Entity;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public class LogService : ILogService
    {
        private TelemetryClient _telemetry;
        protected readonly IMailService _mailService;
        protected readonly IConfigService _configService;
        public LogService(IMailService mailService, IConfigService configService)
        {
            _telemetry = new TelemetryClient();
            _mailService = mailService;
            _configService = configService;
        }
        public void Flush()
        {
            _telemetry.Flush();
        }
        /// <summary>
        /// Write the log to send grid or application insight.
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public async Task WriteLog(LogEntity entity)
        {
            var properties = new Dictionary<string, string>();
            properties.Add("Subject", entity.Subject);
            properties.Add("Message", entity.Message);
            properties.Add("Log ID", entity.LogId);
            properties.Add("Action", entity.Action);
            properties.Add("Point Type", entity.PointType);
            if (entity.ActionType == ActionTypeEnum.ErrorLog)
            {
                properties.Add("Detail", entity.Detail);
            }
            properties.Add("Action Type", entity.ActionType.ToString());

            //var metric = new Dictionary<string, double>();
            _telemetry.TrackEvent(string.Format("{0} {1}", entity.PointType, entity.ActionType), properties);

            if (entity.ActionType == ActionTypeEnum.ErrorLog)
            {
                await _mailService.SendPlainTextMail(
                    _configService.SendGridMessageFromAddress,
                    _configService.SendGridMessageFromDisplayName,
                    _configService.SendGridMessageToAddress,
                    $"{entity.LogId} - {entity.Subject}",
                    $"{entity.Subject} - \t{entity.Message} - \t{entity.Detail}");
            }
#if DEBUG
            Flush();
#endif
        }
    }
}
