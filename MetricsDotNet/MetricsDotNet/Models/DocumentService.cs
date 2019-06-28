using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using MetricsDotNet.ViewModels;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Text;

namespace MetricsDotNet.Models
{
    public class DocumentService : IDocumentService
    {
        public void SendInfoToService(FieldsCapturedViewModel fieldsCapturedVM)
        {
            try
            {

                List<Metric> metrics = new List<Metric>();
                if (fieldsCapturedVM.TalkPercentage > 0)
                {
                    metrics.Add(

                    new Metric
                    {
                        metric = "Talk",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.TalkPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.GrowPercentage > 0)
                {

                    metrics.Add(new Metric
                    {
                        metric = "Grow",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.GrowPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }

                if (fieldsCapturedVM.FeedbackPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Feedback",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.FeedbackPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.UtilizationPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Utilization",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.UtilizationPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.FeedbackPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Certified Technical Interviewers",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.CITPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.SucessASMTPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Successful ASMT",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.SucessASMTPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.BandMixPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Band Mix",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.BandMixPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.MgmtVelocityPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Management Velocity",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.MgmtVelocityPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.AttrittionPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Attrition",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.AttrittionPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.RotationAgreePercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Rotation Agreements",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.RotationAgreePercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.RMBandgesPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "RM Badges",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.RMBandgesPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }
                if (fieldsCapturedVM.MentoringPPPercentage > 0)
                {
                    metrics.Add(new Metric
                    {
                        metric = "Mentoring program participation",
                        Measurements = new List<Measurement> { new Measurement { Value = fieldsCapturedVM.MentoringPPPercentage, Date = DateTime.Now.Date.ToString() } }
                    });
                }

                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://netmetrics.azurewebsites.net/");
                    var response = client.PostAsJsonAsync("api/PostMetrics", metrics).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        Console.Write("Success");
                    }
                    else
                    {
                        Console.Write("Error");
                    }
                }



            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public FieldsCapturedViewModel UploadExcelDocument(string excelPath)
        {
            try
            {
                FileInfo file = new FileInfo(excelPath);
                List<KeyValuePair<int, string>> empUnitList = new List<KeyValuePair<int, string>>();
                decimal factorGrow = 0;
                decimal factorTalk = 0;
                decimal factorBench = 0;
                // hoja unit snapshot
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    StringBuilder sb = new StringBuilder();
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[4];
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;
                    var celText = string.Empty;
                    // obtiene empleados con EmpUnit -
                    for (int r = 11; r <= rowCount; r++)
                    {
                        celText = worksheet.Cells[r, 7].Value.ToString();
                        if (celText == "-")
                        {
                            empUnitList.Add(new KeyValuePair<int, string>(r, "Denied"));
                        }
                        else
                        {
                            empUnitList.Add(new KeyValuePair<int, string>(r, "Access"));
                        }
                        Console.WriteLine(celText);
                    }
                    // obtiene empleados factor Grow ( 1 y 2 son permitidos )- GROW
                    int rFactorG = 11;
                    foreach (var item in empUnitList)
                    {
                        celText = worksheet.Cells[rFactorG, 36].Value.ToString();
                        if (item.Value != "Denied")
                        {
                            if (celText.Contains("Last Month"))
                            {
                                factorGrow++;
                            }
                        }
                        rFactorG++;
                    }
                    // obtiene empleados factor Talk 
                    int rFactorT = 11;
                    foreach (var item in empUnitList)
                    {
                        celText = worksheet.Cells[rFactorT, 49].Value.ToString();
                        if (item.Value != "Denied")
                        {
                            if (celText.Contains("<1 m"))
                            {
                                factorTalk++;
                            }
                        }
                        rFactorT++;
                    }
                    // obtiene empleados factor Bench / utilitation 
                    int rFactorB = 11;
                    foreach (var item in empUnitList)
                    {
                        celText = worksheet.Cells[rFactorB, 57].Value.ToString();
                        if (item.Value != "Denied")
                        {
                            if (celText.Contains("No"))
                            {
                                factorBench++;
                            }
                        }
                        rFactorB++;
                    }
                }
                int total = empUnitList.FindAll(k => k.Value.Contains("Access")).Count;
                // bandmix
                float bandMixA2, bandMixA3, bandMixA4;
                using (ExcelPackage packageSeniority = new ExcelPackage(file))
                {
                    StringBuilder sb = new StringBuilder();
                    ExcelWorksheet worksheetSeniority = packageSeniority.Workbook.Worksheets[9];
                    float.TryParse(worksheetSeniority.Cells[10, 3].Value.ToString(), out bandMixA2);
                    float.TryParse(worksheetSeniority.Cells[11, 3].Value.ToString(), out bandMixA3);
                    float.TryParse(worksheetSeniority.Cells[12, 3].Value.ToString(), out bandMixA4);
                }
                var returnVB = new FieldsCapturedViewModel
                {
                    Total = total,
                    TalkCompliance = factorTalk,
                    TalkPercentage = (factorTalk / total) * 100,
                    GrowCompliance = (total - factorGrow),
                    GrowPercentage = ((total - factorGrow) / total) * 100,
                    UtilizationCompliance = factorBench,
                    UtilizationPercentage = (factorBench / total) * 100,
                    BandMixA2 = bandMixA2,
                    BandMixA3 = bandMixA3,
                    BandMixA4 = bandMixA4,
                    BandMixA2Total = (bandMixA2 * 2),
                    BandMixA3Total = (bandMixA3 * 3),
                    BandMixA4Total = (bandMixA4 * 4),
                    BandMixPercentage = (decimal)((bandMixA2 * 2) + (bandMixA3 * 3) + (bandMixA4 * 4)) / total,
                    AttrittionCompliance = 1,
                };
                returnVB.AttrittionPercentage = (decimal)1 / total;
                returnVB.AttrittionPercentage = (decimal)returnVB.AttrittionPercentage * 100;

                DeleteFile(excelPath);

                return returnVB;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public bool DeleteFile(string excelPath)
        {
            try
            {
                File.Delete(excelPath);

                return true;
            }
            catch (Exception e)
            {
                return false;
                throw new Exception(e.Message);
            }

        }

    }

    public class Metric
    {
        public string metric { get; set; }
        public List<Measurement> Measurements { get; set; }
    }

    public class Measurement
    {
        private string _date;
        private decimal _value;

        public string Date
        {
            get
            {
                return _date;
            }
            set
            {
                _date = DateTimeOffset.Now.ToUnixTimeSeconds().ToString();
            }
        }

        public decimal Value
        {
            get
            {
                return _value;
            }
            set
            {
                _value = decimal.Round(value, 2);
            }

        }
    }
}