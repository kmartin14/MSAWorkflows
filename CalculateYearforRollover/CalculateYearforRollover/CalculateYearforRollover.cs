using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;

namespace CalculateYearforRollover
{
    public class CalculateyearforRollover : CodeActivity
    {
        [Input("Year")]
        public InArgument<string> Year { get; set; }

        [Output("Result")]
        public OutArgument<string> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            String valYear = Year.Get<string>(context);
            int intYear = Convert.ToInt32(valYear);
            intYear = intYear + 1;
            string stringYear = Convert.ToString(intYear);            

            Result.Set(context, stringYear);
        }
    }
}
