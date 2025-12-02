using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Documentupdate
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString;
            string excelFilePath = ConfigurationManager.AppSettings["ExcelFilePath"];
            string logFilePath = ConfigurationManager.AppSettings["LogFilePath"];


            
            var candidates = ReadExcelFile(excelFilePath);

          
            foreach (var candidate in candidates)
            {
               
                Loggerfile.LogInfo($"Processing candidate {candidate.ApplicantNo}, indent number {candidate.IndentNo}");

               
                if (IsCandidateValid(connectionString, candidate))
                {
                    
                    string query = string.Format("UPDATE hc_req_resume SET stage_id = {0}, status_id = {1} WHERE applicant_no = {2} AND indent_no = {3}",
                        candidate.UpdatedStageId, candidate.UpdatedStatusId, candidate.ApplicantNo, candidate.IndentNo);

                  
                    ExecuteQuery(connectionString, query);

                  
                    Loggerfile.LogInfo($"Successfully updated candidate {candidate.ApplicantNo} with indent number {candidate.IndentNo}");
                    Console.WriteLine($"Successfully updated candidate {candidate.ApplicantNo} with indent number {candidate.IndentNo}");
                
                }
                else
                {
                   
                    Loggerfile.LogError($"Candidate {candidate.ApplicantNo} with indent number {candidate.IndentNo} failed validation checks.");
                    Console.WriteLine($"Candidate {candidate.ApplicantNo} with indent number {candidate.IndentNo} failed validation checks.");
                }
            }
        }

       
        static List<CandidateData> ReadExcelFile(string filePath)
        {
            try
            {
                DataTable dt = new DataTable();

                using (var workbook = new XLWorkbook(filePath))
                {
                    var ws = workbook.Worksheet(1);

                    // Build columns from header row
                    foreach (var cell in ws.Row(1).CellsUsed())
                        dt.Columns.Add(cell.GetValue<string>());

                    // Fill rows
                    foreach (var row in ws.RowsUsed().Skip(1))
                    {
                        var dr = dt.NewRow();
                        for (int i = 0; i < dt.Columns.Count; i++)
                            dr[i] = row.Cell(i + 1).GetValue<string>();

                        dt.Rows.Add(dr);
                    }
                }

                // Convert DataTable -> List<CandidateData>
                return dt.AsEnumerable()
                         .Select(r => new CandidateData
                         {
                             ApplicantNo = r[0].ToString(),
                             IndentNo = r[1].ToString(),
                             CurrentStageId = r[2].ToString(),
                             CurrentStatusId = r[3].ToString(),
                             UpdatedStageId = r[4].ToString(),
                             UpdatedStatusId = r[5].ToString()
                         })
                         .ToList();
            }
            catch (Exception ex)
            {
                Loggerfile.LogInfo($"Error reading Excel file: {ex.Message}");
                return new List<CandidateData>();
            }
        }


        static bool IsCandidateValid(string connectionString, CandidateData candidate)
        {

            string query = $@"
                   SELECT COUNT(*)
                  FROM hc_resume_bank rb
                 JOIN hc_req_resume rr 
                   ON rr.resid = rb.rid
                JOIN hc_requisitions req
               ON req.rid = rr.reqid
               WHERE rb.candidateno = '{candidate.ApplicantNo}'
             AND req.reqnumber = '{candidate.IndentNo}'
              AND rr.stageid = '{candidate.CurrentStageId}'
                AND rr.statusid = '{candidate.CurrentStatusId}'
                                                             ";



            int count = ExecuteScalarQuery(connectionString, query);

           
            return count > 0;
        }

        
        static int ExecuteScalarQuery(string connectionString, string query)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                return Convert.ToInt32(command.ExecuteScalar());
            }
        }


        static void ExecuteQuery(string connectionString, string query)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                command.ExecuteNonQuery();
            }
        }
    }

    // Candidate Data Class
    public class CandidateData
    {
        public string ApplicantNo { get; set; }
        public string IndentNo { get; set; }
        public string CurrentStageId { get; set; }
        public string CurrentStatusId { get; set; }
        public string UpdatedStageId { get; set; }
        public string UpdatedStatusId { get; set; }
    }

    // Loggerfile Class


    
}
