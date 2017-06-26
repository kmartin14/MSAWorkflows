using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow; 

namespace CalculateYearEndDate
{
    public class CalculateYearEndDate : CodeActivity 
    {
        [Input("Year")]
        public InArgument<string> Year { get; set; }

        [Input("Fiscal Year End")]
        [AttributeTarget("new_clientid","new_fiscalyearend")]
        public InArgument<OptionSetValue> FiscalYearEnd { get; set; }

        [Output("Year End Date")]
        public OutArgument<DateTime> YearEndDate { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            String valYear = Year.Get<string>(context);
            OptionSetValue valFiscalYearEnd = FiscalYearEnd.Get<OptionSetValue>(context);
            DateTime retYearEndDate = new DateTime(); 

            int intYear = Convert.ToInt32(valYear);

            if (valFiscalYearEnd.Value < 13 && valFiscalYearEnd.Value > 0)
            {
                //build a date with our month and year for the first of the month
                retYearEndDate = new DateTime(intYear, valFiscalYearEnd.Value, 1);

                //add month
                retYearEndDate = retYearEndDate.AddMonths(1);

                //subtract a day 
                retYearEndDate = retYearEndDate.AddDays(-1);
            }

            YearEndDate.Set(context, retYearEndDate);
        }
    }
}
