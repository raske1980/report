using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Report.Renderers
{
    /// <summary>
    /// 
    /// </summary>
    public abstract class RendererBase : IRenderer
    {
        public abstract byte[] Render(Base.Report report);

        public static string ColorToRgb(System.Drawing.Color color)
        {
            return color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
        }

    }
}
