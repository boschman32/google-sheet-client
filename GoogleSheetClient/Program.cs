using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Sheets.v4.Data;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

using CommandLine;

namespace GoogleSheetClient
{
    class Options
    {
        [Option('i', "spreadsheet-id", Required = true, HelpText = "The Id of the Google spreadsheet")]
        public string SpreadsheetId { get; set; }

        [Option('o', "output-dir", Default = @"output")]
        public string OutputDir { get; set; }

        [Option('f', "file-names", Required = true, HelpText = "Names to give the range json files. Example: -f database database1")]
        public IEnumerable<string> FileNames { get; set; }

        [Option('r', "data-ranges", Required = true, HelpText = "Ranges in the sheet, Example: -r Tab!A1:E5 \"Tab 2!A1:E5\"")]
        public IEnumerable<string> DataRanges { get; set; }

        [Option('h', "header-count", HelpText = "The number of headers in the sheets, Example: 1, 2, 4. \nNote: by default all sheets have 1 as header count each entry in the command line corresponds to the sheet in the dataranges.")]
        public IEnumerable<int> HeaderCounts { get; set; }

        [Option('k', "keep-open", Default = false, HelpText = "Whether you want to keep the window open after it is done.")]
        public bool KeepOpenAfterCompletion { get; set; }

        [Option('d', "debug", Default = false, HelpText = "Debug, ask for input if there is an error.")]
        public bool Debug { get; set; }
    }

    class Program
    {
        private static readonly bool ForceDebug = false;

        private static readonly string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        private static readonly string ApplicationName = "GoogleSheetClient";
        private static readonly string SheetObjectRegex = "\\((obj)\\)";
        private static readonly string SheetListRegex = "\\((list)\\)";

        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(RunWithOptions)
                .WithNotParsed(errs =>
                {
                    foreach (var error in errs)
                    {
                        Console.Error.WriteLine(error.ToString());
                    }

                    if (ForceDebug)
                    {
                        Console.WriteLine("Press any key to exit...");
                        Console.Read();
                    }

                    Environment.ExitCode = 0;
                });
        }

        private static void RunWithOptions(Options opts)
        {
            if (opts.DataRanges.Count() != opts.FileNames.Count())
            {
                Console.Error.WriteLine("Number of data-ranges given and number of given file-names are not equal!");
                if (opts.KeepOpenAfterCompletion || ForceDebug)
                {
                    Console.WriteLine("Press any key to close...");
                    Console.Read();
                }
                return;
            }

            using var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read);
            // The file token.json stores the user's access and refresh tokens, and is created
            // automatically when the authorization flow completes for the first time.
            UserCredential credential = null;
            try
            {
                Console.WriteLine("Accessing google login token...");
                CancellationTokenSource cts = new CancellationTokenSource(new TimeSpan(0, 0, 30));
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    cts.Token,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);

                if (opts.Debug || ForceDebug)
                {
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadLine();
                }

                Environment.ExitCode = 1;
            }

            if (credential != null)
            {
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });

                var data = GetSheetData(service, opts.SpreadsheetId, opts.DataRanges.ToList(), opts.HeaderCounts.ToList(), opts.Debug);

                if (data != null)
                {
                    var outputDir = opts.OutputDir;
                    var fileNames = opts.FileNames.ToList();

                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    int filenameIndex = 0;
                    foreach (var d in data)
                    {
                        var filename = fileNames[filenameIndex];
                        var output = Path.Combine(outputDir, filename) + ".json";
                        using StreamWriter file = File.CreateText(output);
                        using JsonTextWriter writer = new JsonTextWriter(file) { Formatting = Formatting.Indented };
                        d.WriteTo(writer);

                        filenameIndex++;
                    }

                    Console.WriteLine("Successfully exported sheet json data...");
                }
            }

            if (opts.KeepOpenAfterCompletion || ForceDebug)
            {
                Console.WriteLine("Press any key to close...");
                Console.Read();
            }
        }

        private static List<JArray> GetSheetData(SheetsService service, string spreadsheetId, List<string> dataRanges, List<int> headerDepth, bool Debug)
        {
            BatchGetValuesResponse response = null;
            try
            {
                SpreadsheetsResource.ValuesResource.BatchGetRequest request =
                    service.Spreadsheets.Values.BatchGet(spreadsheetId);
                request.Ranges = dataRanges;

                response = request.Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                if (Debug || ForceDebug)
                {
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadLine();
                }

                Environment.ExitCode = 2;
            }

            if (response != null)
            {
                List<JArray> data = new List<JArray>();
                IList<ValueRange> valueRanges = response.ValueRanges;

                using (var rangeBar = new ProgressBar("Ranges: "))
                {
                    if (valueRanges != null && valueRanges.Count > 0)
                    {
                        for (int i = 0; i < valueRanges.Count; i++)
                        {
                            string range = valueRanges[i].Range;
                            List<List<object>> values = CreateGridData(valueRanges[i].Values);
                            if (values != null && values.Count > 1)
                            {
                                data.Add(ParseValuesToJsonData(range, values, (i < headerDepth.Count ? headerDepth[i] : 1)));
                            }
                            rangeBar.Report((double)i / valueRanges.Count);
                        }
                        return data;
                    }
                }
            }

            Console.WriteLine("No data found!");
            return null;
        }

        private static List<List<object>> CreateGridData(IList<IList<object>> data)
        {
            List<List<object>> gridData = new List<List<object>>();

            int gridWith = 0;
            //First find the longest data list inside the given data.
            foreach (var d in data)
            {
                gridWith = Math.Max(gridWith, d.Count);
            }

            for (int row = 0; row < data.Count; row++)
            {
                gridData.Add(data[row].ToList());

                if (gridData[row].Count >= gridWith)
                {
                    continue;
                }

                //If the amount of data isn't the maximum add some empty data in their to form a grid.
                int diff = gridWith - data[row].Count;

                List<string> emptyEntries = new List<string>();
                for (int i = 0; i < diff; i++)
                {
                    emptyEntries.Add("");
                }
                gridData[row].AddRange(emptyEntries);
            }

            return gridData;
        }

        private static JArray ParseValuesToJsonData(string sheetRange, List<List<object>> data, int headerDepth)
        {
            var jsonData = new JArray();
            Console.WriteLine("Parsing sheet: ({0})...", sheetRange);

            using (var progressBar = new ProgressBar(sheetRange))
            {
                for (int r = headerDepth; r < data.Count; r++)
                {
                    JObject jsonEntry = new JObject();

                    bool newEntryAdded = false;
                    for (int c = 0; c < data[r].Count; c++)
                    {
                        var valueObj = data[r][c];

                        if (!string.IsNullOrWhiteSpace(valueObj.ToString()) && !string.IsNullOrWhiteSpace(data[r][0].ToString()))
                        {
                            //Grab the top most header of the sheet.
                            var topHeader = data[0][c].ToString();
                            //Get a clean header without white spaces and macros: (obj), (list), etc.
                            var cleanHeader = GetCleanHeader(topHeader);

                            int tempRow = r; //Don't skip any sheet rows that are new entries in the list. 
                            JToken value = GetJsonValue(topHeader, cleanHeader, valueObj, jsonEntry, ref data, headerDepth, ref tempRow, ref c);

                            if (value == null)
                            {
                                continue;
                            }

                            if (!string.IsNullOrWhiteSpace(topHeader))
                            {
                                jsonEntry.Add(cleanHeader, value);
                            }
                            else
                            {
                                jsonData.Add(value);
                            }
                        }

                        progressBar.Report((r / data.Count) + (c / data[r].Count));
                    }

                    if (jsonEntry.HasValues)
                    {
                        jsonData.Add(jsonEntry);
                    }
                }
            }
            Console.WriteLine("Done parsing sheet: ({0})!", sheetRange);
            return jsonData;
        }

        //Create a grid of all the possible items that can be contained in the google sheet list.
        private static List<List<object>> CreateSubGridData(ref List<List<object>> data, int headerDepth, int row, int column, int maxEntries = -1)
        {
            List<List<object>> subGrid = new List<List<object>>();

            var dataTopHeaders = data[0];
            int numEntryValues = 0;
            int numEntries = 0;

            //Find any values inside the grid
            for (int r = row; r < data.Count; r++)
            {
                //Skip the first column check since that contains the id (corresponded row data of the list).
                //If the column contains something it is the end of the list.
                if (r > row && !string.IsNullOrWhiteSpace(data[r][0].ToString()))
                {
                    break;
                }

                numEntryValues = 0;
                subGrid.Add(new List<
                    object>());
                for (int c = column; c < data[row].Count; c++)
                {
                    //Skip the first top header since that is always the name of the list.
                    //If the next column contains something it means we know how many entry values the grid has.
                    if (c > column && !string.IsNullOrWhiteSpace(dataTopHeaders[c].ToString()))
                    {
                        break;
                    }

                    //Always add the row add the end list.
                    subGrid[^1].Add(data[r][c]);
                    numEntryValues++;
                }

                numEntries++;
                if (maxEntries != -1 && numEntries >= maxEntries)
                {
                    break;
                }
            }

            //Parse any headers that are part of the sub grid.
            var topHeaders = new List<List<object>>();
            for (int i = 1; i < headerDepth; i++)
            {
                topHeaders.Add(new List<object>());
                for (int j = column; j < column + numEntryValues; j++)
                {
                    //Always add data to the last list index.
                    topHeaders[^1].Add(data[i][j]);
                }
            }

            subGrid.InsertRange(0, topHeaders);

            return subGrid;
        }

        private static JObject CreateJsonObject(ref List<List<object>> data, int headerDepth, ref int currentRow, ref int currentColumn)
        {
            JObject obj = new JObject();

            //Start parsing data after the headers.
            for (int r = headerDepth - 1; r < data.Count; ++r)
            {
                for (int c = 0; c < data[r].Count; c++)
                {
                    //Add a new value to the list entry.
                    var header = data[0][c].ToString();
                    //Get a clean header without white spaces and macros: (obj), (list), etc.
                    var cleanHeader = GetCleanHeader(header);

                    JToken value = GetJsonValue(header, cleanHeader, data[r][c], obj, ref data, headerDepth - 1, ref r, ref c);

                    if (value != null)
                    {
                        obj.Add(cleanHeader, value);
                    }
                }
            }

            //Tell the loop that is looping over all the data that we have processed already parts of that data since it was a list so we can skip that amount.
            currentColumn += data[0].Count - 1;
            currentRow += data.Count - headerDepth;

            return obj;
        }

        //From a given grid list we make a json array containing all the items and values with them.
        private static JArray CreateJsonArray(ref List<List<object>> data, int headerDepth, ref int currentRow, ref int currentColumn)
        {
            JArray list = new JArray();

            for (int r = headerDepth - 1; r < data.Count; r++)
            {
                JObject entry = new JObject();
                for (int c = 0; c < data[r].Count; c++)
                {
                    //Add a new value to the list entry.
                    var header = data[0][c].ToString();
                    //Get a clean header without white spaces and macros: (obj), (list), etc.
                    var cleanHeader = GetCleanHeader(header);
                    JToken value = GetJsonValue(header, cleanHeader, data[r][c], entry, ref data, headerDepth  - 1, ref r, ref c);

                    if (value == null)
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(header))
                    {
                        entry.Add(cleanHeader, value);
                    }
                    else
                    {
                        list.Add(value);
                    }
                }

                //If the entry has any values we can add it to the list.
                if (entry.HasValues)
                {
                    list.Add(entry);
                }
            }

            //Tell the loop that is looping over all the data that we have processed already parts of that data since it was a list so we can skip that amount.1
            currentColumn += data[0].Count - 1;
            currentRow += data.Count - headerDepth;

            return list;
        }

        private static JToken GetJsonValue(string header, string cleanHeader, object value, JObject entry, ref List<List<object>> data, int headerDepth, ref int r, ref int c)
        {
            //First check what we are dealing with (list, object, or plain value).
            JToken valueToken = null;
            if (IsList(header))
            {
                List<List<object>> subGrid = CreateSubGridData(ref data, headerDepth, r, c);
                if (subGrid.Count > 0)
                {
                    valueToken = CreateJsonArray(ref subGrid, headerDepth, ref r, ref c);
                }
            }
            else if (IsObject(header))
            {
                List<List<object>> subGrid = CreateSubGridData(ref data, headerDepth, r, c, 1);
                if (subGrid.Count > 0)
                {
                    valueToken = CreateJsonObject(ref subGrid, headerDepth, ref r, ref c);
                }
            }
            else
            {
                valueToken = GetObjectValue(value);
            }

            //Check if the token contains any values.
            if (IsJTokenNullOrEmpty(valueToken))
            {
                return null;
            }

            //if the header isn't empty and there is already a entry with that key.
            if (!string.IsNullOrWhiteSpace(cleanHeader) && entry.ContainsKey(cleanHeader))
            {
                Console.WriteLine("[List addition] Json entry already exists: ({0}), couldn't add list: ({1})...", cleanHeader, valueToken);
                return null;
            }

            return valueToken;
        }

        //Get the right JSON value from the data.
        private static JValue GetObjectValue(object data)
        {
            if (int.TryParse(data.ToString(), out var integer))
            {
                return new JValue(integer);
            }

            if (float.TryParse(data.ToString(), NumberStyles.Float, CultureInfo.CreateSpecificCulture("en-US"), out var floating))
            {
                return new JValue(floating);
            }

            if (bool.TryParse(data.ToString(), out var boolean))
            {
                return new JValue(boolean);
            }

            return (JValue)JToken.FromObject(data);
        }

        private static bool IsObject(string header)
        {
            return Regex.IsMatch(header, SheetObjectRegex);
        }

        private static bool IsList(string header)
        {
            return Regex.IsMatch(header, SheetListRegex);
        }

        //Remove any whitespace or object/list tags.
        private static string GetCleanHeader(string input)
        {
            string noWhiteSpace = new string(input.ToCharArray()
                .Where(c => !char.IsWhiteSpace(c))
                .ToArray());

            noWhiteSpace = Regex.Replace(noWhiteSpace, SheetListRegex, "");
            noWhiteSpace = Regex.Replace(noWhiteSpace, SheetObjectRegex, "");

            return noWhiteSpace;
        }

        private static bool IsJTokenNullOrEmpty(JToken token)
        {
            return (token == null) ||
                   (token.Type == JTokenType.Array && !token.HasValues) ||
                   (token.Type == JTokenType.Object && !token.HasValues) ||
                   (token.Type == JTokenType.String && token.ToString() == string.Empty) ||
                   (token.Type == JTokenType.Null);
        }

        private static bool IsJTokenAContainer(JToken token)
        {
            return (token != null && 
                    ((token.Type == JTokenType.Array && !token.HasValues) || 
                     (token.Type == JTokenType.Object && !token.HasValues)));
        }
    }
}
