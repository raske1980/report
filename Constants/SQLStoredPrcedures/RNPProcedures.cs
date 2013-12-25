using System;
using System.Collections.Generic;
using System.Text;

namespace Constants.SQLStoredPrcedures
{
    public static class RNPProcedures
    {

        public static string GetRNPBase
        {
            get 
            {
                return
                    @"Select 
                        rnpRow.Name as RowName
                        , subDiv.Name as DivName
                        , rtDisc.CreditECTS as CreditECTS
                        , rnpRow.HourFact as HourFact
                        , (select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 24) as ExamenCount
                        , (select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 25) as ZalikCount
                        , (select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 37) as MkrCount
                        ,  (select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 35) as CourseProjectCount
                        , (select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 34) as CourseWorkCount
                        , ISNULL((select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 40),0) + ISNULL((select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 39),0) as RgrCount
                        , (select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 36) as DkrCount	
                        , (select sum(CountControlTest) from dbo.cRNPLoad l where l.cRNPRowId = rnpRow.cRNPRowId and l.DcLoadId = 38) as ReferatCount
                    from dbo.cRNPRow rnpRow
                    inner join 	dbo.DcSubdivision subDiv on rnpRow.DcSubdivisionId = subDiv.DcSubdivisionId
                    inner join dbo.RtDiscipline rtDisc on rtDisc.RtDisciplineId = rnpRow.RtDisciplineId";
            }
        }
    }
}
