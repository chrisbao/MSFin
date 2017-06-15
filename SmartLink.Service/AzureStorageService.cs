using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Microsoft.WindowsAzure.Storage.Table;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public sealed class AzureStorageService : IAzureStorageService
    {
        private object _queuelock = new object();
        private object _tablelock = new object();

        private readonly IConfigService _configService;

        private CloudStorageAccount _cloudStorageAccount;
        private CloudQueueClient _queueClient;
        private CloudTableClient _tableClient;
        private IDictionary<string, CloudQueue> _cloudQueues;
        private IDictionary<string, CloudTable> _cloudTables;
        public CloudQueueClient QueueClient
        {
            get
            {
                if (_queueClient == null)
                {
                    _queueClient = _cloudStorageAccount.CreateCloudQueueClient();
                }
                return _queueClient;
            }
        }

        public CloudTableClient TableClient
        {
            get
            {
                if (_tableClient == null)
                {
                    _tableClient = _cloudStorageAccount.CreateCloudTableClient();
                }
                return _tableClient;
            }
        }

        public AzureStorageService(IConfigService configService)
        {
            _configService = configService;
            var connectionString = _configService.AzureWebJobsStorage;
            _cloudStorageAccount = Microsoft.WindowsAzure.Storage.CloudStorageAccount.Parse(connectionString);
            ServicePointManager.FindServicePoint(_cloudStorageAccount.QueueEndpoint).UseNagleAlgorithm = false;
            ServicePointManager.FindServicePoint(_cloudStorageAccount.TableEndpoint).UseNagleAlgorithm = false;
            _cloudQueues = new Dictionary<string, CloudQueue>();
            _cloudTables = new Dictionary<string, CloudTable>();
        }
        /// <summary>
        /// Get cloud queue by name
        /// </summary>
        /// <param name="queueName"></param>
        /// <returns></returns>
        public CloudQueue GetQueue(string queueName)
        {
            lock (_queuelock)
            {
                if (!_cloudQueues.ContainsKey(queueName))
                {
                    var queue = QueueClient.GetQueueReference(queueName);
                    queue.CreateIfNotExists();
                    _cloudQueues.Add(queueName, queue);
                }
            }
            return _cloudQueues[queueName];
        }
        /// <summary>
        /// Get cloud table by name
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public CloudTable GetTable(string tableName)
        {
            lock (_tablelock)
            {
                if (!_cloudTables.ContainsKey(tableName))
                {
                    var table = TableClient.GetTableReference(tableName);
                    table.CreateIfNotExists();
                    _cloudTables.Add(tableName, table);
                }
            }
            return _cloudTables[tableName];
        }
          /// <summary>
          /// Add message to the queue.
          /// </summary>
          /// <param name="queueMessage"></param>
          /// <param name="queueName"></param>
          /// <returns></returns>
        public Task WriteMessageToQueue(string queueMessage, string queueName)
        {
            var queue = GetQueue(queueName);
            return queue.AddMessageAsync(new CloudQueueMessage(queueMessage));
        }
    }
}