using System;

namespace Report
{
    public interface IRenderer
    {
        byte[] Render(Report.Base.Report report);
    }
}
