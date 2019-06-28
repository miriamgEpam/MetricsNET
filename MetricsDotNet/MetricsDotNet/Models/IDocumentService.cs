using MetricsDotNet.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetricsDotNet.Models
{
    public interface IDocumentService
    {
        FieldsCapturedViewModel UploadExcelDocument(string route);
        void SendInfoToService(FieldsCapturedViewModel fieldsCapturedVM);
    }
}
