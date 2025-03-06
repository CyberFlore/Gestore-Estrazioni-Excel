using System;
using System.Collections.Generic;

namespace ExcelTemplateManager
{
    [Serializable]
    public class Template
    {
        public List<ColumnInfo> Columns { get; set; }
        
        public Template()
        {
            Columns = new List<ColumnInfo>();
        }
    }
}
