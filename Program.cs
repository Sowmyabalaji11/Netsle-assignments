using System.Data;
using ExcelDataReader;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using ClosedXML.Excel;

namespace EmployeeDataComparison
{
    class Program
    {
        static void Main(string[] args){
            try
            {
                // Create a new DataTable
                DataTable excelData = new DataTable();
                // Load the workbook
                using (XLWorkbook workbook = new XLWorkbook(@"book1.xlsx"))
                {
                    // Get the first worksheet
                    var worksheet = workbook.Worksheet(1);
                    // Loop through the columns in the worksheet and add them to the DataTable
                    bool firstRow = true;
                    foreach (IXLRow row in worksheet.Rows())
                    {
                        if (firstRow)
                        {
                            // Assuming the first row contains the column headers
                            foreach (IXLCell cell in row.Cells())
                            {
                                excelData.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            // Add the rest of the data to the DataTable
                            excelData.Rows.Add();
                            int i = 0;
                            foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                            {
                                excelData.Rows[excelData.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }
                    }
                }
                Console.WriteLine(excelData);
                // Connect to Active Directory
                using (PrincipalContext context = new PrincipalContext(ContextType.Domain,"DomainName"))
                {
                    // Process data in batches for performance optimization
                    int batchSize = 10;
                    int rowCount = excelData.Rows.Count;
                    List<Task> tasks = new List<Task>();

                    for (int i = 0; i < rowCount; i += batchSize)
                    {
                        DataTable batchData = GetBatchData(excelData, i, batchSize);
                        tasks.Add(ProcessBatchData(context, batchData));
                    }

                    // Wait for all tasks to complete
                    Task.WhenAll(tasks).Wait();
                }
                // Save the updated Excel data
                SaveExcelData(excelData, "book1.xlsx");

                Console.WriteLine("Employee data comparison and update completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                // Log the error to a file or external logging service
            }
        }

        static DataTable LoadExcelData(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet().Tables[0];
                }
            }
        }

        static DataTable GetBatchData(DataTable sourceData, int startIndex, int batchSize)
        {
            DataTable batchData = sourceData.Clone();
            for (int i = startIndex; i < Math.Min(startIndex + batchSize, sourceData.Rows.Count); i++)
            {
                batchData.ImportRow(sourceData.Rows[i]);
            }
            return batchData;
        }

        static async Task ProcessBatchData(PrincipalContext context, DataTable batchData)
        {
            await Task.Run(() =>
            {
                // Add a new column to store the comparison result
                batchData.Columns.Add("ManagerNameComparisonResult", typeof(string));

                foreach (DataRow row in batchData.Rows)
                {
                    try
                    {
                        string employeeId = row["EmployeeId"].ToString();
                        string managerIdExcel = row["ManagerId"].ToString();

                        UserPrincipal user = FindUserByCustomAttribute(context, "EmployeeId", employeeId);

                        if (user != null)
                        {
                            string managerIdAD = GetManagerIdFromAD(user);
                            string managerNameAD = GetManagerNameFromAD(managerIdAD, context);

                            if (!string.IsNullOrEmpty(managerIdAD))
                            {
                                string managerIdADExtracted = managerIdAD.Split(',')[0].Split('=')[1];
                                if (!managerIdExcel.Equals(managerIdADExtracted, StringComparison.OrdinalIgnoreCase))
                                {
                                    row["ManagerIdUpdated"] = managerIdADExtracted;
                                    row["ManagerNameUpdated"] = managerNameAD;
                                    // Compare manager names and set the comparison result
                                    row["ManagerNameComparisonResult"] = row["ManagerName"].ToString().Equals(managerNameAD, StringComparison.OrdinalIgnoreCase) ? "Correct" : "Wrong";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing row: {ex.Message}");
                        // Log the error to a file or external logging service
                    }
                }
            });
        }

        static UserPrincipal FindUserByCustomAttribute(PrincipalContext context, string attributeName, string attributeValue)
        {
            PrincipalSearcher searcher = new PrincipalSearcher();
            UserPrincipal user = new UserPrincipal(context);
            user.EmployeeId = attributeValue;
            searcher.QueryFilter = user;
            return searcher.FindOne() as UserPrincipal;
        }

        static string GetManagerIdFromAD(UserPrincipal user)
        {
            DirectoryEntry userEntry = (DirectoryEntry)user.GetUnderlyingObject();
            string managerDN = userEntry.Properties["manager"].Value?.ToString();

            if (!string.IsNullOrEmpty(managerDN))
            {
                return managerDN.Split(',')[0].Split('=')[1];
            }

            return string.Empty;
        }

        static string GetManagerNameFromAD(string managerId, PrincipalContext context)
        {
            UserPrincipal manager = FindUserByCustomAttribute(context, "EmployeeId", managerId);
            return manager?.DisplayName ?? string.Empty;
        }

        static void SaveExcelData(DataTable data, string filePath)
        {
            using (var stream = File.Create(filePath))
            {
                using (var writer = new StreamWriter(stream))
                {
                    foreach (DataRow row in data.Rows)
                    {
                        writer.WriteLine(string.Join(",", row.ItemArray));
                    }
                }
            }
        }
    }
}
