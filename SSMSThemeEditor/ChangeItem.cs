using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSMSThemeEditor
{
    class ChangeItem
    {
        public IntPtr LabelPtr { get; set; }
        public string Style { get; set; }
        public Color OldLabelColor { get; set; }
        public Color NewLabelColor { get; set; }

        public override string ToString()
        {
            return $"{Style} -> {OldLabelColor.ToString()} TO {NewLabelColor.ToString()}";
        }
    }
}
