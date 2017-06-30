/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Caching;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public class DocumentService : IDocumentService
    {
        private readonly IConfigService _configService;

        public DocumentService(IConfigService configService)
        {
            _configService = configService;
        }

        protected string ClientID { get { return _configService.WebJobClientId; } }

        protected string Authority
        {
            get
            {
                return _configService.AzureADInstance + _configService.AzureADTenantId;
            }
        }

        protected string Resource
        {
            get
            {
                return _configService.SharePointUrl;
            }
        }

        protected string CertificatedPassword
        {
            get
            {
                return _configService.CertificatePassword;
            }
        }

        protected ObjectCache LocalCache
        {
            get
            {
                return MemoryCache.Default;
            }
        }

        /// <summary>
        /// Update the bookmark value by destination points in the specific word file.
        /// </summary>
        /// <param name="destinationFileName"></param>
        /// <param name="destinationPoints"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public async Task<DocumentUpdateResult> UpdateBookmarkValueAsync(string destinationFileName, IEnumerable<DestinationPoint> destinationPoints, string value)
        {
            DocumentUpdateResult retValue = new DocumentUpdateResult() { IsSuccess = true };
            try
            {
                string authHeader = GetAuthorizationHeader();
                FileContextInfo fileContextInfo = await GetFileContextInfoAsync(authHeader, destinationFileName);
                var stream = await GetFileStreamAsync(authHeader, destinationFileName, fileContextInfo);
                stream.Seek(0, SeekOrigin.Begin);
                UpdateStream(destinationPoints, value, stream, retValue);
                await UploadStreamAsync(authHeader, destinationFileName, stream, fileContextInfo);
                return retValue;
            }
            catch (Exception ex)
            {
                retValue.IsSuccess = false;
                retValue.Message.Add(ex.ToString());
            }
            return retValue;
        }

        /// <summary>
        /// Upload the updated word file to SP document library where the word file exsited.
        /// </summary>
        /// <param name="authHeader"></param>
        /// <param name="destinationFileName"></param>
        /// <param name="stream"></param>
        /// <param name="contextInfo"></param>
        /// <returns></returns>
        /// <summary>
        /// Create the authorization header to authenticate with SP site.
        /// </summary>
        /// <returns></returns>
        private string GetAuthorizationHeader()
        {
            string authorizationHeader = LocalCache["AuthorizationHeader"] as string;
            lock (typeof(DocumentService))
            {

                if (string.IsNullOrWhiteSpace(authorizationHeader))
                {
                    AuthenticationContext authenticationContext = new AuthenticationContext(Authority, false);
                    var cert = GetCertificate();

                    //Try to use http://stackoverflow.com/questions/6392268/x509certificate-keyset-does-not-exist
                    using (cert.GetRSAPrivateKey()) { }

                    ClientAssertionCertificate cac = new ClientAssertionCertificate(ClientID, cert);

                    var authenticationResult = authenticationContext.AcquireTokenAsync(Resource, cac).Result;
                    authorizationHeader = authenticationResult.CreateAuthorizationHeader();
                    LocalCache.Set("AuthorizationHeader", authorizationHeader, authenticationResult.ExpiresOn.AddMinutes(-2));
                }
            }
            return authorizationHeader;
        }

        /// <summary>
        /// Get the certificate via a specific path.
        /// </summary>
        /// <returns></returns>
        private X509Certificate2 GetCertificate()
        {
            //read the certificate private key from the executing location
            //NOTE: This is a hack…Azure Key Vault is best approach
            var certPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            certPath = certPath.Substring(0, certPath.LastIndexOf('\\')) + "\\cert\\" + _configService.CertificateFile;
            var certfile = System.IO.File.OpenRead(certPath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            return new X509Certificate2(
                certificateBytes,
                CertificatedPassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);
        }

        static public void UpdateStream(IEnumerable<DestinationPoint> destinationPoints, string value, Stream stream, DocumentUpdateResult updateResult)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                foreach (var destinationPoint in destinationPoints)
                {
                    int foundTimes = 0;
                    var position = destinationPoint.RangeId;
                    var formats = destinationPoint.CustomFormats.ToList();
                    var formattedValue = GetFormattedValue(value, formats);

                    //Update value in sdtBlock
                    foundTimes += UpdateValueInSdtBlock(formattedValue, wordDoc, position);
                    //Update value in sdtCell
                    foundTimes += UpdateValueInSdtCell(formattedValue, wordDoc, position);
                    //Update value in sdtRun
                    foundTimes += UpdateValueInSdtRun(formattedValue, wordDoc, position);
                    updateResult.Message.Add($"Tag:{position} found {foundTimes}(updated to {formattedValue}).");
                }

                wordDoc.MainDocumentPart.Document.Save();
            }
        }

        static private async Task UploadStreamAsync(string authHeader, string destinationFileName, Stream stream, FileContextInfo contextInfo)
        {
            var fileAbsolutePath = new Uri(destinationFileName).AbsolutePath;
            int index = fileAbsolutePath.LastIndexOf('/');
            var filePath = fileAbsolutePath.Substring(0, index);
            var fileName = fileAbsolutePath.Substring(index + 1);

            using (var client = new WebClient())
            {
                client.Headers.Add(HttpRequestHeader.Authorization, authHeader);
                client.Headers.Add(HttpRequestHeader.Accept, "application/json");

                var uploadUri = new Uri($"{contextInfo.WebUrl}/_api/web/GetFolderByServerRelativeUrl(@filePath)/Files/add(url='{fileName}',overwrite=true)?@filePath='{filePath}'");
                var returnValue = await client.UploadDataTaskAsync(uploadUri, "POST", (stream as MemoryStream).ToArray());
            }
        }

        /// <summary>
        /// Update the destination points value in word file.
        /// </summary>
        /// <param name="destinationPoints"></param>
        /// <param name="value"></param>
        /// <param name="stream"></param>
        /// <param name="updateResult"></param>
        /// <summary>
        /// Update the destination point value under sdtrun node by position (tag).
        /// </summary>
        /// <param name="value"></param>
        /// <param name="wordDoc"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        static private int UpdateValueInSdtRun(string value, WordprocessingDocument wordDoc, string position)
        {
            int retValue = 0;
            var items = wordDoc.MainDocumentPart.Document.Descendants<SdtRun>().Where(
                o =>
                {
                    var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                    if (tagedItem != null)
                    {
                        return tagedItem.Val == position;
                    }
                    return false;
                });
            foreach (var item in items)
            {
                var texts = item.Descendants<Text>();
                var textCount = texts.Count();
                if (textCount == 1)
                {
                    texts.First().Text = value;
                    retValue++;
                }
                else if (textCount > 1)
                {
                    texts.First().Text = value;
                    for (int i = 1; i < textCount; i++)
                    {
                        texts.ElementAt(i).Text = string.Empty;
                    }
                    /*
                    //remove all paragraphs from the content cell
                    item.SdtContentRun.RemoveAllChildren<Paragraph>();
                    //create a new paragraph containing a run and a text element
                    Paragraph newParagraph = new Paragraph();
                    Run newRun = new Run();
                    Text newText = new Text(value);
                    newRun.Append(newText);
                    newParagraph.Append(newRun);
                    item.Append(newParagraph);
                    */
                    retValue++;
                }
            }
            return retValue;
        }

        /// <summary>
        /// Update the destination point value under SDTBlock node by position (tag).
        /// </summary>
        /// <param name="value"></param>
        /// <param name="wordDoc"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        static private int UpdateValueInSdtBlock(string value, WordprocessingDocument wordDoc, string position)
        {
            int retValue = 0;
            var items = wordDoc.MainDocumentPart.Document.Descendants<SdtBlock>().Where(
                o =>
                {
                    var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                    if (tagedItem != null)
                    {
                        return tagedItem.Val == position;
                    }
                    return false;
                });
            foreach (var item in items)
            {
                var texts = item.Descendants<Text>();
                var textCount = texts.Count();
                if (textCount == 1)
                {
                    texts.First().Text = value;
                    retValue++;
                }
                else if (textCount > 1)
                {
                    texts.First().Text = value;
                    for (int i = 1; i < textCount; i++)
                    {
                        texts.ElementAt(i).Text = string.Empty;
                    }
                    /*
                    //remove all paragraphs from the content cell
                    item.SdtContentBlock.RemoveAllChildren<Paragraph>();
                    //create a new paragraph containing a run and a text element
                    Paragraph newParagraph = new Paragraph();
                    Run newRun = new Run();
                    Text newText = new Text(value);
                    newRun.Append(newText);
                    newParagraph.Append(newRun);
                    item.Append(newParagraph);
                    */
                    retValue++;
                }
            }
            return retValue;
        }

        /// <summary>
        ///     /// Update the destination point value under SDTCELL node by position (tag).
        /// </summary>
        /// <param name="value"></param>
        /// <param name="wordDoc"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        static private int UpdateValueInSdtCell(string value, WordprocessingDocument wordDoc, string position)
        {
            int retValue = 0;
            var items = wordDoc.MainDocumentPart.Document.Descendants<SdtCell>().Where(
                o =>
                {
                    var tagedItem = o.SdtProperties.Elements<Tag>().FirstOrDefault();
                    if (tagedItem != null)
                    {
                        return tagedItem.Val == position;
                    }
                    return false;
                });
            foreach (var item in items)
            {
                var tableCell = item.SdtContentCell.Elements<TableCell>().FirstOrDefault();
                if (tableCell != null)
                {
                    var texts = tableCell.Descendants<Text>();
                    var textCount = texts.Count();
                    if (textCount == 1)
                    {
                        texts.First().Text = value;
                        retValue++;
                    }
                    else if (textCount > 1)
                    {
                        texts.First().Text = value;
                        for (int i = 1; i < textCount; i++)
                        {
                            texts.ElementAt(i).Text = string.Empty;
                        }
                        /*
                        //remove all paragraphs from the content cell
                        tableCell.RemoveAllChildren<Paragraph>();
                        //create a new paragraph containing a run and a text element
                        Paragraph newParagraph = new Paragraph();
                        Run newRun = new Run();
                        Text newText = new Text(value);
                        newRun.Append(newText);
                        newParagraph.Append(newRun);
                        tableCell.Append(newParagraph);
                        */
                        retValue++;
                    }
                }
            }
            return retValue;
        }

        /// <summary>
        /// Get the word file stream by file name.
        /// </summary>
        /// <param name="authHeader"></param>
        /// <param name="destinationFileName"></param>
        /// <param name="fileContextInfo"></param>
        /// <returns></returns>
        static private async Task<Stream> GetFileStreamAsync(string authHeader, string destinationFileName, FileContextInfo fileContextInfo)
        {
            var filePath = new Uri(destinationFileName).AbsolutePath;
            var stream = new System.IO.MemoryStream();
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers.Add(HttpRequestHeader.Authorization, authHeader);
                    var fileUri = new Uri(destinationFileName);
                    var downloadData = await client.DownloadDataTaskAsync($"{fileContextInfo.WebUrl}/_api/web/getfilebyserverrelativeurl('{filePath}')/$value");
                    await stream.WriteAsync(downloadData, 0, downloadData.Length);
                    return stream;
                }
            }
            catch (Exception ex)
            {
                stream.Dispose();
                throw ex;
            }
        }

        /// <summary>
        /// Get file context information by file absolute URL.
        /// </summary>
        /// <param name="authHeader"></param>
        /// <param name="destinationFileName"></param>
        /// <returns></returns>
        static private async Task<FileContextInfo> GetFileContextInfoAsync(string authHeader, string destinationFileName)
        {
            var fileFolder = destinationFileName.Substring(0, destinationFileName.LastIndexOf('/'));
            using (var client = new WebClient())
            {
                client.Headers.Add(HttpRequestHeader.Authorization, authHeader);
                client.Headers.Add(HttpRequestHeader.Accept, "application/json");

                var fileInfo = await client.UploadStringTaskAsync($"{fileFolder}/_api/contextinfo", "POST", "");

                var jFileInfo = JObject.Parse(fileInfo);
                return new FileContextInfo()
                {
                    FormDigest = jFileInfo["FormDigestValue"].Value<string>(),
                    WebUrl = jFileInfo["WebFullUrl"].Value<string>()
                };
            }
        }

        /// <summary>
        /// Convert the value to the formatted value.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="formats"></param>
        /// <returns></returns>
        static private string GetFormattedValue(string value, List<CustomFormat> formats)
        {
            var originalValue = value;
            var hasDollar = originalValue.IndexOf("$") > -1;
            var hasComma = originalValue.IndexOf(",") > -1;
            var hasPercent = originalValue.IndexOf("%") > -1;
            var valueWithOutSymbol = originalValue.Replace("$", "").Replace(",", "");
            var isNumeric = Regex.IsMatch(valueWithOutSymbol, @"^[+-]?\d*[.]?\d*$");
            var isDate = IsDate(originalValue);

            foreach (var format in formats)
            {
                if (format.Name == "ConvertToHundreds")
                {
                    if (isNumeric)
                    {
                        value = (Convert.ToDecimal(valueWithOutSymbol) / 100).ToString();
                        if (hasComma)
                        {
                            value = AddComma(value);
                        }
                        if (hasDollar)
                        {
                            value = "$" + value;
                        }
                    }
                }
                else if (format.Name == "ConvertToThousands")
                {
                    if (isNumeric)
                    {
                        value = (Convert.ToDecimal(valueWithOutSymbol) / 1000).ToString();
                        if (hasComma)
                        {
                            value = AddComma(value);
                        }
                        if (hasDollar)
                        {
                            value = "$" + value;
                        }
                    }
                }
                else if (format.Name == "ConvertToMillions")
                {
                    if (isNumeric)
                    {
                        value = (Convert.ToDecimal(valueWithOutSymbol) / 1000000).ToString();
                        if (hasComma)
                        {
                            value = AddComma(value);
                        }
                        if (hasDollar)
                        {
                            value = "$" + value;
                        }
                    }
                }
                else if (format.Name == "ConvertToBillions")
                {
                    if (isNumeric)
                    {
                        value = (Convert.ToDecimal(valueWithOutSymbol) / 1000000000).ToString();
                        if (hasComma)
                        {
                            value = AddComma(value);
                        }
                        if (hasDollar)
                        {
                            value = "$" + value;
                        }
                    }
                }
                else if (format.Name == "AddDecimalPlace")
                {
                    if (isNumeric)
                    {
                        if (value.IndexOf(".") > -1)
                        {
                            value = value + "0";
                        }
                        else
                        {
                            value = value + ".0";
                        }
                    }
                }
                else if (format.Name == "ShowNegativesAsPositives")
                {
                    var tempValue = value.Replace("$", "").Replace("-", "").Replace("%", "").Replace("(", "").Replace(")", "");
                    if (Regex.IsMatch(tempValue, @"^[+-]?\d*[.]?\d*$"))
                    {
                        value = tempValue;
                        if (hasPercent)
                        {
                            value = value + "%";
                        }
                        if (hasDollar)
                        {
                            value = "$" + value;
                        }
                    }
                }
                else if (format.Name == "IncludeThousandDescriptor")
                {
                    if (isNumeric)
                    {
                        value = value + " thousand";
                    }
                }
                else if (format.Name == "IncludeMillionDescriptor")
                {
                    if (isNumeric)
                    {
                        value = value + " million";
                    }
                }
                else if (format.Name == "IncludeBillionDescriptor")
                {
                    if (isNumeric)
                    {
                        value = value + " billion";
                    }
                }
                else if (format.Name == "IncludeDollarSymbol")
                {
                    if (value.IndexOf("$") == -1)
                    {
                        value = "$" + value;
                    }
                }
                else if (format.Name == "ExcludeDollarSymbol")
                {
                    if (value.IndexOf("$") > -1)
                    {
                        value = value.Replace("$", "");
                    }
                }
                else if (format.Name == "DateShowLongDateFormat")
                {
                    if (isDate)
                    {
                        var date = Convert.ToDateTime(value);
                        value = GetMonth(date.Month - 1) + " " + date.Day + ", " + date.Year;
                    }
                }
                else if (format.Name == "DateShowYearOnly")
                {
                    if (isDate)
                    {
                        var date = Convert.ToDateTime(value);
                        value = date.Year.ToString();
                    }
                }
                else if (format.Name == "ConvertNegativeSymbolToParenthesis")
                {
                    if (value.IndexOf("-") > -1)
                    {
                        var h = value.IndexOf("$") > -1;
                        var tempValue = value.Replace("$", "").Replace("-", "").Replace("(", "").Replace(")", "");
                        value = "(" + tempValue + ")";
                        if (h)
                        {
                            value = "$" + value;
                        }
                    }
                }
            }
            return value;
        }

        static private bool IsDate(string value)
        {
            var flag = false;
            try
            {
                var date = Convert.ToDateTime(value);
                flag = date.Year > 0;
            }
            catch
            {
                flag = false;
            }
            return flag;
        }

        static private string AddComma(string value)
        {
            string newstr = string.Empty;
            Regex r = new Regex(@"(\d+?)(\d{3})*(\.\d+|$)");
            Match m = r.Match(value);
            newstr += m.Groups[1].Value;
            for (int i = 0; i < m.Groups[2].Captures.Count; i++)
            {
                newstr += "," + m.Groups[2].Captures[i].Value;
            }
            newstr += m.Groups[3].Value;
            return newstr;
        }

        static private string GetMonth(int month)
        {
            var _m = new string[12] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            return _m[month];
        }
    }
}