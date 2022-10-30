using System;

namespace MasterConverter
{
    public class CyberneCsv : ICsvConvertible
    {
        [Ignore]
        public string header { get; private set; } = "\"REGION_IDENTFY_CD\",\"CYBERNE_CD\",\"END_YMD_HMS\",\"STATION_NAME_8\",\"STATION_NAME_4\",\"STATION_NAME_YOMI\",\"COMPANY_CD\",\"COMPANY_SYMBOL\"";

        [EndComma]
        [DoubleQuotes]
        public string AreaIdentityCode { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string CyberneCode { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string ExpiryDate { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string Name8Letters { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string Name4Letters { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string Name10Letters { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string CompanyCode { get; set; }

        [DoubleQuotes]
        public string CompanyName { get; set; }
    }
}
