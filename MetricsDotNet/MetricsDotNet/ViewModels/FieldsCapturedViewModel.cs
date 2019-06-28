using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace MetricsDotNet.ViewModels
{
    public class FieldsCapturedViewModel
    {
        public int  Total { get; set; }
        
        public decimal TalkCompliance { get; set; }
        
        public decimal TalkPercentage{ get; set; }

        public decimal GrowCompliance { get; set; }
        public decimal GrowPercentage { get; set; }

        public float FeedbackCompliance { get; set; }
        public decimal FeedbackPercentage { get; set; }

        public decimal UtilizationCompliance { get; set; }
        public decimal UtilizationPercentage { get; set; }

        public float CITCompliance { get; set; }
        public decimal CITPercentage { get; set; }

        public float SucessASMTCompliance { get; set; }
        public decimal SucessASMTPercentage { get; set; }


        public float BandMixA2 { get; set; }
        public float BandMixA3 { get; set; }
        public float BandMixA4 { get; set; }
        public float BandMixA2Total { get; set; }
        public float BandMixA3Total { get; set; }
        public float BandMixA4Total { get; set; }
        public float BandMixCompliance { get; set; }
        public decimal BandMixPercentage { get; set;  }

        public float MgmtVelocityCompliance { get; set; }
        public decimal MgmtVelocityPercentage { get; set; }

        public float AttrittionCompliance { get; set; }
        public decimal AttrittionPercentage { get; set; }

        public float RotationAgreeCompliance { get; set; }
        public decimal RotationAgreePercentage { get; set; }

        public float RMBadgesCompliance { get; set; }
        public decimal RMBandgesPercentage { get; set; }

        public float MentoringPPCompliance { get; set; }
        public decimal MentoringPPPercentage { get; set; }
    }


}
