using System.Collections.Generic;

namespace ToExcel
{
    public sealed class AgencyObject
    {
        public AgencyMetadata Metadata { get; set; }
        public List<AgencyRateStructure> RateStructures { get; set; }
        public AgencyList CapacityCharge { get; set; }
    }
}
