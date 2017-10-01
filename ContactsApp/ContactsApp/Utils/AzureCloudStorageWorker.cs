using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;

namespace ContactsApp.Utils
{
    public class AzureCloudStorageWorker
    {
        private CloudTable _table;
        private CloudStorageAccount _storageAccount;

        internal async Task<bool> ConnectAsync(string connectionString, string tableName)
        {
            if (!CloudStorageAccount.TryParse(connectionString, out _storageAccount))
            {
                return false;
            }

            _table = _storageAccount.CreateCloudTableClient().GetTableReference(tableName);
            await _table.CreateIfNotExistsAsync();
            return true;
        }

        internal async Task InsertOrReplace(List<ITableEntity> entityList)
        {
            if (entityList == null)
            {
                throw new ArgumentNullException(nameof(entityList));
            }

            foreach (var entity in entityList)
            {
                await _table.ExecuteAsync(TableOperation.InsertOrReplace(entity));
            }
        }
    }
}
