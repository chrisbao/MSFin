using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public interface IMailService
    {
        Task SendPlainTextMail(string fromAddress, string fromDisplayName, IEnumerable<string> toAddresses, string subject, string content);
    }
}
