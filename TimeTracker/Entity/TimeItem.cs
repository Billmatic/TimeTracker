using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimeTracker.Entity
{
    public class TimeItem
    {
        public string title { get; set; }

        public string time { get; set; }

        public string description { get; set; }

        public bool? isBillable { get; set; }

        public bool? isExternalComment { get; set; }

        public bool? isCRMSubmitted { get; set; }

        public int? index { get; set; }

        public Guid? crmTaskId { get; set; }

        public Guid? crmExternalCommentId { get; set; }

        public string project { get; set; }
    }
}
