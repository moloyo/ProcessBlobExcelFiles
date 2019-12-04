
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
namespace ProcessExcelFiles
{
    public static class Durable
    {
        [FunctionName("Durable")]
        public static async Task<List<string>> RunOrchestrator(
            [OrchestrationTrigger] IDurableOrchestrationContext context)
        {
            var outputs = new List<string>();
            var file = context.GetInput<File>();            

            outputs.AddRange(await Task.WhenAll(file.BreakRowsIntoStacks().Select(j => context.CallActivityAsync<string>("Process_Rows", new FileJob { Rows = j, Dictionary = file.Dictionary }))));

            await context.CallActivityAsync("LogProcessRows", outputs);

            return outputs;

        }

        [FunctionName("LogProcessRows")]
        public async static Task LogProcessRows([ActivityTrigger] IDurableActivityContext input)
        {
            var logs = input.GetInput<IEnumerable<string>>();
            var instanceId = input.InstanceId;

            var connString = Environment.GetEnvironmentVariable("ConnectionString");
            // ConfigurationManager.ConnectionStrings[0].ConnectionString;
            using (var conn = new SqlConnection(connString))
            {
                conn.Open();
                var values = string.Join(", ", logs.Select(l => $"('{l}', '{instanceId}', '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}')"));
                var query = $"INSERT INTO Logs VALUES {values}";
                using (var cmd = new SqlCommand(query, conn))
                {
                    await cmd.ExecuteNonQueryAsync();
                }
            }
        }

        [FunctionName("Process_Rows")]
        public async static Task<string> ProcessRows([ActivityTrigger] IDurableActivityContext input, ILogger log)
        {
            var obj = input.GetInput<FileJob>();
            var rows = obj.Rows;
            var tasks = new List<Task>();
            foreach (var row in rows)
            {
                var translatedCells = row.Cells.Select(c => c.FileType == FileType.SharedString ? obj.Dictionary[int.Parse(c.Value ?? "0")] : c.Value).ToArray();
                tasks.Add(PersistRow(translatedCells));
            }
            await Task.WhenAll(tasks);
            log.LogInformation($"{rows.Count()} were processed.");
            return $"{rows.Count()} were processed.";
        }

        private async static Task PersistRow(string[] translatedCells)
        {
            var connString = Environment.GetEnvironmentVariable("ConnectionString");
            using (var conn = new SqlConnection(connString))
            {
                conn.Open();
                var query = $"INSERT INTO Rows VALUES ('{translatedCells[0]}', '{translatedCells[1]}')";
                using (var cmd = new SqlCommand(query, conn))
                {
                    await cmd.ExecuteNonQueryAsync();
                }
            }
        }

        [FunctionName("Durable_HttpStart")]
        public static async Task<IActionResult> HttpStart(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")]HttpRequest req,
            [DurableClient]IDurableOrchestrationClient starter,
            ILogger log)
        {
            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            string fileName = req.Query["fileName"];
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            fileName = fileName ?? data?.fileName ?? "Client Sample Data file.xlsx";

            var file = await RecoverFile(fileName);

            var instanceId = await starter.StartNewAsync("Durable", null, file);
            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");
            return starter.CreateCheckStatusResponse(req, instanceId);
        }

        private async static Task<File> RecoverFile(string fileName)
        {
            ICollection<FileRow> rows = new List<FileRow>();
            var dictionary = new Dictionary<int, string>();

            var connectionString = Environment.GetEnvironmentVariable("BlobStorage");
            var account = CloudStorageAccount.Parse(connectionString);
            var blobClient = account.CreateCloudBlobClient();
            var containerReference = blobClient.GetContainerReference("imported-blobs");
            var blob = await containerReference.GetBlobReferenceFromServerAsync(fileName);
            using (var stream = new MemoryStream())
            {
                await blob.DownloadToStreamAsync(stream);

                using (var doc = SpreadsheetDocument.Open(stream, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var reader = OpenXmlReader.Create(worksheetPart);
                    // Recover dictionary
                    dictionary = workbookPart.GetPartsOfType<SharedStringTablePart>()
                        .First()
                        .SharedStringTable
                        .ChildElements
                        .Select((value, index) => new { value, index })
                        .ToDictionary((e) => e.index, (e) => e.value.InnerText);

                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Row))
                        {
                            var row = reader.LoadCurrentElement() as Row;
                            rows.Add(new FileRow { Cells = row.Descendants<Cell>().Select(c => new FileCell() { Value = c.CellValue?.InnerText, FileType = c.DataType != null ? (FileType?)((int)c.DataType?.Value) : null }).ToList() });
                        }
                    };
                }
            }

            return new File(rows, dictionary);
        }
    }
}