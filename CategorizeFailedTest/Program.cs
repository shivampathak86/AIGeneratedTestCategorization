using System.Net;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using OfficeOpenXml;

class Program
{
    public class Category
    {
        public string? CategoryName { get; set; }
        public List<string>? TestCases { get; set; }
    }

    static void Main(string[] args)
    {
        Approach3();
    }
    //static void Approach1()
    //{
    //    string apiKey = "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"; // Replace your API key
    //    string endpoint = "https://api.openai.com/v1/chat/completions";
    //    string filePath = "C:\\Repos\\misc\\Trx2Excel\\Trx2Excel\\bin\\Debug\\test2.xlsx";
    //    string sheetName = "TestResult";

    //    List<string> failureMessages = new List<string>();
    //    List<string> testCaseNames = new List<string>();
    //   // List<string> testCaseNamespaces = new List<string>();

    //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    //    using (var package = new ExcelPackage(new FileInfo(filePath)))
    //    {
    //        var worksheet = package.Workbook.Worksheets[sheetName];
    //        int rowCount = worksheet.Dimension.Rows;

    //        for (int row = 1; row <= rowCount; row++) // assuming the data starts from row 1
    //        {
    //            string? failureMessage = worksheet.Cells[row, 4].Value?.ToString();
    //            string? testCaseName = worksheet.Cells[row, 1].Value?.ToString();
    //            //string? testCaseNamespace = inputWorksheet.Cells[row, 3].Value?.ToString();

    //            if (!string.IsNullOrEmpty(failureMessage) && !string.IsNullOrEmpty(testCaseName) ) //!string.IsNullOrEmpty(testCaseNamespace))
    //            {
    //                failureMessages.Add(failureMessage);
    //                testCaseNames.Add(testCaseName);
    //                //testCaseNamespaces.Add(testCaseNamespace);
    //            }
    //        }
    //    }
    //    string finalPrompt = string.Empty;
    //    int maxMessages = 20; // Choose an appropriate number of messages to include in the conversation
    //    int startIndex = Math.Max(0, failureMessages.Count - maxMessages);
    //    for (int i = startIndex; i < failureMessages.Count; i++)
    //    {
    //        string failureMessage = failureMessages[i];
    //        string testCaseName = testCaseNames[i];
    //        //Build Prompt
    //        finalPrompt += $"Failure Message: {failureMessage}\nTest Case Name: {testCaseName}\n";
    //    }

    //    List<Category> categories =  GenerateCategoryUsingAI(apiKey, endpoint, finalPrompt);
    //    Console.WriteLine("Failure Category\t\tTest Cases");
    //    Console.WriteLine("-----------------------------------------------");
    //    foreach (var category in categories)
    //    {
    //        Console.WriteLine($"{category.CategoryName}\t{string.Join(", ", category.TestCases)}");
    //    }
    //}

    //static void Approach2()
    //{
    //    string apiKey = "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"; // Replace your API key
    //    string endpoint = "https://api.openai.com/v1/chat/completions";
    //    string filePath = "C:\\Repos\\misc\\Trx2Excel\\Trx2Excel\\bin\\Debug\\test.xlsx";
    //    string sheetName = "TestResult";

    //    List<string> failureMessages = new List<string>();
    //    List<string> testCaseNames = new List<string>();
    //    List<string> testCaseNamespaces = new List<string>();

    //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    //    using (var package = new ExcelPackage(new FileInfo(filePath)))
    //    {
    //        var worksheet = package.Workbook.Worksheets[sheetName];
    //        int rowCount = worksheet.Dimension.Rows;

    //        for (int row = 1; row <= rowCount; row++) // assuming the data starts from row 1
    //        {
    //            string? failureMessage = worksheet.Cells[row, 4].Value?.ToString();
    //            string? testCaseName = worksheet.Cells[row, 1].Value?.ToString();
    //            string? testCaseNamespace = worksheet.Cells[row, 3].Value?.ToString();

    //            if (!string.IsNullOrEmpty(failureMessage) && !string.IsNullOrEmpty(testCaseName) && !string.IsNullOrEmpty(testCaseNamespace))
    //            {
    //                failureMessages.Add(failureMessage);
    //                testCaseNames.Add(testCaseName);
    //                testCaseNamespaces.Add(testCaseNamespace);
    //            }
    //        }
    //    }

    //    for (int i = 1; i < failureMessages.Count; i++)
    //    {
    //        string failureMessage = failureMessages[i];
    //        string testCaseName = testCaseNames[i];
    //        string testCaseNamespace = testCaseNamespaces[i];

    //        // Generate failureCategory using AI model
    //        string modelInput = $"Failure Message: {failureMessage}\nTest Case Name: {testCaseName}\n";

    //        List<Category> categories = GenerateCategoryUsingAI(apiKey, endpoint, modelInput);
    //        Console.WriteLine("Failure Category\t\tTest Cases");
    //        Console.WriteLine("-----------------------------------------------");
    //        foreach (var category in categories)
    //        {
    //            Console.WriteLine($"{category.CategoryName}\t{string.Join(", ", category.TestCases)}");
    //        }

    //        Thread.Sleep(20000);
    //    }
    //}

    static void Approach3()
    {
        string apiKey = "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"; // Replace your API key
        string endpoint = "https://api.openai.com/v1/chat/completions";
        string filePath = "C:\\Repos\\misc\\Trx2Excel\\Trx2Excel\\bin\\Debug\\test2.xlsx";
        string sheetName = "TestResult";

        List<string> failureMessages = new List<string>();
        List<string> testCaseNames = new List<string>();
        // List<string> testCaseNamespaces = new List<string>();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var inputWorksheet = package.Workbook.Worksheets[sheetName];
            int rowCount = inputWorksheet.Dimension.Rows;

            for (int row = 1; row <= rowCount; row++) // assuming the data starts from row 1
            {
                string? failureMessage = inputWorksheet.Cells[row, 4].Value?.ToString();
                string? testCaseName = inputWorksheet.Cells[row, 1].Value?.ToString();
                //string? testCaseNamespace = inputWorksheet.Cells[row, 3].Value?.ToString();

                if (!string.IsNullOrEmpty(failureMessage) && !string.IsNullOrEmpty(testCaseName)) //!string.IsNullOrEmpty(testCaseNamespace))
                {
                    failureMessages.Add(failureMessage);
                    testCaseNames.Add(testCaseName);
                    //testCaseNamespaces.Add(testCaseNamespace);
                }
            }
        }
 
        int chunkSize = 25; // Choose the number of rows per chunk
        var failureCategories = new List<Category>();
        for (int start = 0; start < failureMessages.Count; start += chunkSize)
        {
            int end = Math.Min(start + chunkSize, failureMessages.Count);
            List<string> chunkFailureMessages = failureMessages.GetRange(start, end - start);
            List<string> chunkTestCaseNames = testCaseNames.GetRange(start, end - start);

            string prompt = string.Empty;

            for (int i = 1; i < chunkFailureMessages.Count; i++)
            {
                string failureMessage = chunkFailureMessages[i];
                string testCaseName = chunkTestCaseNames[i];
                // Build Prompt
                prompt += $"\nFailure Message: {failureMessage}\nTest Case Name: {testCaseName}\n";
            }

            List<Category> categories = GenerateCategoryUsingAI(apiKey, endpoint, prompt);

            failureCategories.AddRange(categories);
        }
        string finalPrompt = string.Empty;
        foreach (var failureCategory in failureCategories)
        {
            finalPrompt += $"\nFailure Message: {failureCategory.CategoryName}\nTest Case Name: {string.Join(",",failureCategory.TestCases)}\n";
        }
        failureCategories.Clear();
        failureCategories = GenerateCategoryUsingAI(apiKey, endpoint, finalPrompt);

        // Create Excel package
        using (var excelPackage = new ExcelPackage())
        {
            var outputWorksheet = excelPackage.Workbook.Worksheets.Add("Failure Categories");

            // Add table headers
            outputWorksheet.Cells[1, 1].Value = "Failure Category";
            outputWorksheet.Cells[1, 2].Value = "Test Cases";

            // Populate table data
            int rowIndex = 2;
            foreach (var failureCategory in failureCategories)
            {
                outputWorksheet.Cells[rowIndex, 1].Value = failureCategory.CategoryName;
                outputWorksheet.Cells[rowIndex, 2].Value = string.Join(", ", failureCategory.TestCases);
                rowIndex++;
            }

            // Auto-fit column widths
            outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();

            // Save the Excel file
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{DateTime.Now}_FailureCategories.xlsx");
            var file = new FileInfo(path);
            excelPackage.SaveAs(file);
        }
    }
    static List<Category> GenerateCategoryUsingAI(string apiKey, string endpoint, string modelInput)
    {
        using (HttpClient httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            var requestData = new
            {
                model = "gpt-3.5-turbo",
                messages = new[]
             {      
                  new {role = "system", content ="Create failure categories based on the failure messages, and group test cases that have similar or correlated failures under each category.Ensure that similar failure categories are grouped together, even if there is some overlap in failure category values. Please keep output format exactly \nFailure Category: \nTest Case Name: \n What Test Case Do: .There should not be none in category"},
                   new {role="user", content = modelInput}
              }
            };
            var requestDataJson = JsonConvert.SerializeObject(requestData);
            var content = new StringContent(requestDataJson, System.Text.Encoding.UTF8, "application/json");
            var response = httpClient.PostAsync(endpoint, content).Result;
            var responseContent = response.Content.ReadAsStringAsync().Result;

            if (response.StatusCode == HttpStatusCode.OK)
            {
            
                var jsonResponse = JsonConvert.DeserializeObject<dynamic>(responseContent);
                var categories = new List<Category>();

                foreach (var choice in jsonResponse.choices)
                {
                    var result = choice.message.content.ToString();
                    var categoryTestPairs = ExtractCategoryTestPairs(result);
                    foreach (var pair in categoryTestPairs)
                    {
                        var category = pair.Key;
                        var testCases = pair.Value;
                        categories.Add(new Category { CategoryName = category, TestCases = testCases });
                    }
                }

                return categories;
            }
            else
            {
                return null;
            }
        }
    }
   
    static Dictionary<string, List<string>> ExtractCategoryTestPairs(string modelOutput)
    {
        var pairs = new Dictionary<string, List<string>>();
        var lines = modelOutput.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        string currentCategory = "";
        List<string> currentTestCases = new List<string>();

        foreach (var line in lines)
        {
            if (line.Trim().StartsWith("Failure Category: "))
            {
                if (!string.IsNullOrEmpty(currentCategory))
                {
                    pairs[currentCategory] = currentTestCases;
                    currentTestCases = new List<string>();
                }

                currentCategory = line.Substring("Failure Category: ".Length).Trim();
            }
            else if (line.Trim().StartsWith("Test Case Name: "))
            {
                var testCase = line.Trim().Substring("Test Case Name: ".Length).Trim();
                currentTestCases.Add(testCase);
            }
        }

        if (!string.IsNullOrEmpty(currentCategory))
        {
            pairs[currentCategory] = currentTestCases;
        }

        return pairs;
    }
}

