using System;
using System.Collections.Generic;

namespace ExcelTemplateManager
{
    [Serializable]
    public class ColumnInfo
    {
        public string Header { get; set; }
        public string OriginalHeader { get; set; }  
        public int OriginalIndex { get; set; }
        public List<string> ValidationOptions { get; set; }
        public bool HasValidation { get; set; }

        public ColumnInfo()
        {
            Header = string.Empty;
            OriginalHeader = string.Empty;
            OriginalIndex = -1;
            ValidationOptions = new List<string>();
            HasValidation = false;
        }
    }
}