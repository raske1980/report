using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RNPExcelExport.Data.Item
{
    public class RNPBase
    {
        public String RowName
        {
            get;
            set;
        }

        public String DivName
        {
            get;
            set;
        }

        public Double CreditECTS
        {
            get;
            set;
        }

        public Int32 HourFact
        {
            get;
            set;
        }

        public Int16? ExamenCount
        {
            get;
            set;
        }

        public Int16? ZalikCount
        {
            get;
            set;
        }

        public Int16? MkrCount
        {
            get;
            set;
        }

        public Int16? CourseProjectCount
        {
            get;
            set;
        }

        public Int16? CourseWorkCount
        {
            get;
            set;
        }

        public Int16 RgrCount
        {
            get;
            set;
        }

        public Int16? DkrCount
        {
            get;
            set;
        }

        public Int16? ReferatCount
        { get; set; }
    }
}


