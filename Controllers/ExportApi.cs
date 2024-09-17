using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ExportApi.ExportModel;
using System;
using Microsoft.Data.SqlClient;
using System.Data;

namespace ExportApi.Controllers
{
  

        [Route("api/[controller]")]
        [ApiController]
        public class ExportApi : ControllerBase
    {
            [HttpPost("export")]
            public async Task<IActionResult> ExportToExcel(ExportModels request)
            {
                if (request == null || string.IsNullOrEmpty(request.TableName))
                {
                    return BadRequest("Table name and connection details are required.");
                }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var connectionString = $"Server={request.Server};Database={request.Database};User Id={request.UserId};Password={request.Password};MultipleActiveResultSets=true;TrustServerCertificate=True;";


                using var connection = new SqlConnection(connectionString);
                await connection.OpenAsync();

                var command = new SqlCommand($"SELECT * FROM [{request.TableName}]", connection);
                var adapter = new SqlDataAdapter(command);
                var dataTable = new DataTable();
                adapter.Fill(dataTable);
                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add(request.TableName);

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{request.TableName}.xlsx");
            }
        }
    }

