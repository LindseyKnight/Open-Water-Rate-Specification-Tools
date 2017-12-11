using System.Collections.Generic;

namespace ToExcel
{
    public sealed class AgencyList
    {
        public List<string> DependsOn { get; set; }
        public Dictionary<string, List<string>> Values { get; set; }
    }
}
