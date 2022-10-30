using System;

namespace MasterConverter
{
    public class PurseCsv : ICsvConvertible
    {
        [Ignore]
        public string header { get; private set; } = "\"IDENTFY_FLG\",\"COMPANY_CD\",\"PURSE_MAX_ADULT\",\"PURSE_MAX_CHILD\",\"PURSE_MAX_ADULT_DISCOUNT\",\"PURSE_MAX_CHILD_DISCOUNT\"";

        [EndComma]
        [DoubleQuotes]
        public string IdentityFlag { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string CompanyCode { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string MaxAdultCharge { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string MaxChildCharge { get; set; }

        [EndComma]
        [DoubleQuotes]
        public string MaxAdultDiscount { get; set; }

        [DoubleQuotes]
        public string MaxChildDiscount { get; set; }

        [Ignore]
        public string Yobi { get; set; }
    }
}
