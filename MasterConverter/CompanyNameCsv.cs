using System;

namespace MasterConverter
{
    public class CompanyNameCsv : ICsvConvertible
    {
        [Ignore]
        public string header { get; private set; } = "\"IDENTFY_FLG\",\"COMPANY_CD\",\"COMPANY_SNAME\"";

        [EndComma]
        [DoubleQuotes]
        public string IdentityFlag { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string CompanyCode { get; set; }

        [DoubleQuotes]
        public string CompanyName { get; set; }
    }
}
