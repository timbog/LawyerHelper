using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LawyerHelper
{
    class mainDataGridItem
    {
        public float number{get;set;}

        public string aim { get; set; }

        public string planDate { get; set; }

        public int timeFromBegin { get; set; }

        public string factDate { get; set; }

        public int pastDue { get; set; }

        public string comment { get; set; }

        public mainDataGridItem(float number, string aim, string planDate, int timeFromBegin, string factDate, int pastDue, string comment)
        {
            this.number = number;
            this.aim = aim;
            this.planDate = planDate;
            this.timeFromBegin = timeFromBegin;
            this.factDate = factDate;
            this.pastDue = pastDue;
            this.comment = comment;
        }
    }
}
