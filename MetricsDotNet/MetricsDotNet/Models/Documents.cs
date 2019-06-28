using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace MetricsDotNet.Models
{
    public class Documents
    {
        [Required]
        public string uploadRoot { get; set; }

    }
}
