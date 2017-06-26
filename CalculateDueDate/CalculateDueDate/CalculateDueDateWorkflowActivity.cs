using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities; 
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;

//Add error handling for blank inputs

namespace CalculateDueDateWorkflowActivity
{
    public class CalculateDueDateWorkflowActivity : CodeActivity 
    {
        [Input("Year End Date")]
        public InArgument<DateTime> YearEndDate { get; set; }

        [Input("Calculation Method")]
        [AttributeTarget("new_taxreturn","new_duedatecalculationmethod")]
        public InArgument<OptionSetValue> CalculationMethod { get; set; }

        [Input("Months Later Due")]
        public InArgument<int> MonthsLaterDue { get; set; }

        [Input("Extension 1 - Months Later Due")]
        public InArgument<int> ext1MonthsLaterDue { get; set; }

        [Input("Extension 2 - Months Later Due")]
        public InArgument<int> ext2MonthsLaterDue { get; set; }
        
        [Input("Day Due")]
        public InArgument<int> DayDue { get; set; }

        [Input("Extension 1 - Day Due")]
        public InArgument<int> ext1DayDue { get; set; }

        [Input("Extension 2 - Day Due")]
        public InArgument<int> ext2DayDue { get; set; }
        
        [Input("Days Later Due")]
        public InArgument<int> DaysLaterDue { get; set; }

        [Input("Extension 1 - Days Later Due")]
        public InArgument<int> ext1DaysLaterDue { get; set; }

        [Input("Extension 2 - Days Later Due")]
        public InArgument<int> ext2DaysLaterDue { get; set; }

        [Input("Times Extended")]
        [AttributeTarget("new_taxformpreparationjob", "new_timesextended")]
        public InArgument<OptionSetValue> TimesExtended { get; set; }

        [Output("Result")]
        public OutArgument<DateTime> result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            DateTime DueDate = YearEndDate.Get<DateTime>(context);
            OptionSetValue valCalculationMethod = CalculationMethod.Get<OptionSetValue>(context);
            int valMonthsLaterDue= MonthsLaterDue.Get<int>(context);
            int valext1MonthsLaterDue = ext1MonthsLaterDue.Get<int>(context);
            int valext2MonthsLaterDue = ext2MonthsLaterDue.Get<int>(context);
            int valDayDue = DayDue.Get<int>(context);
            int valext1DayDue = ext1DayDue.Get<int>(context);
            int valext2DayDue = ext2DayDue.Get<int>(context);
            int valDaysLaterDue = DaysLaterDue.Get<int>(context);
            int valext1DaysLaterDue = ext1DaysLaterDue.Get<int>(context);
            int valext2DaysLaterDue = ext2DaysLaterDue.Get<int>(context);
            OptionSetValue valTimesExtended = TimesExtended.Get<OptionSetValue>(context);
            int valMonth;
            int valYear;
            string calcDate = "";

            DueDate = DueDate.ToLocalTime();

            if (valTimesExtended.Value == 100000000)  //if (valTimesExtended.Equals(100000000))
            {
                //Calculation Method = Months Later - Specific Day of Month for Extension 0
                if (valCalculationMethod.Value == 100000001) //if (valCalculationMethod.Equals(100000001))
                {
                    if (valDayDue <= 0)
                    {
                        valDayDue = 1;
                    }

                    DateTime DueDateAddMonths = DueDate.AddMonths(valMonthsLaterDue);
                    valMonth = DueDateAddMonths.Month;
                    valYear = DueDateAddMonths.Year; 

                    calcDate = DueDateAddMonths.Month + "/" + valDayDue + "/" + DueDateAddMonths.Year;

                    while (DateTime.TryParse(calcDate, out DueDateAddMonths) == false)
                    {
                        valDayDue = valDayDue - 1;
                        calcDate = valMonth + "/" + valDayDue + "/" + valYear;
                    }               

                    DueDate = DueDateAddMonths; 

                }
                //Calculation Method = Months Later for Extension 0
                else if (valCalculationMethod.Value == 100000002)
                {
                    DueDate = DueDate.AddMonths(valMonthsLaterDue);
                }
                //Calculation Method = Month Later - Last Day of Month Extension 0
                else if (valCalculationMethod.Value == 100000003)
                {
                    DateTime DueDateMonthplusone = DueDate.AddMonths(valMonthsLaterDue + 1);
                    DueDateMonthplusone = new DateTime(DueDateMonthplusone.Year, DueDateMonthplusone.Month, 1);
                    DateTime DueDateDayminusone = DueDateMonthplusone.AddDays(-1);
                    DueDate = new DateTime(DueDateDayminusone.Year, DueDateDayminusone.Month, DueDateDayminusone.Day);
                }
                //Calculation Method = Days Later for Extension 0
                else
                {
                    DueDate = DueDate.AddDays(valDaysLaterDue);
                }
            }
            else if (valTimesExtended.Value == 100000001)  //if (valTimesExtended.Equals(100000001))
            {
                //Calculation Method = Months Later - Specific Day of Month Extension 1
                if (valCalculationMethod.Value == 100000001) //if (valCalculationMethod.Equals(100000001))
                {
                    if (valext1DayDue <= 0)
                    {
                        valext1DayDue = 1;
                    }

                    DateTime DueDateAddMonths = DueDate.AddMonths(valext1MonthsLaterDue);
                    valMonth = DueDateAddMonths.Month;
                    valYear = DueDateAddMonths.Year;

                    calcDate = DueDateAddMonths.Month + "/" + valext1DayDue + "/" + DueDateAddMonths.Year;

                    while (DateTime.TryParse(calcDate, out DueDateAddMonths) == false)
                    {
                        valext1DayDue = valext1DayDue - 1;
                        calcDate = valMonth + "/" + valext1DayDue + "/" + valYear;
                    }

                    DueDate = DueDateAddMonths;

                }
                //Calculation Method = Months Later for Extension 1
                else if (valCalculationMethod.Value == 100000002)
                {
                    DueDate = DueDate.AddMonths(valext1MonthsLaterDue);
                }
                //Calculation Method = Month Later - Last Day of Month Extension 1
                else if (valCalculationMethod.Value == 100000003)
                {
                    DateTime DueDateMonthplusone = DueDate.AddMonths(valext1MonthsLaterDue + 1);
                    DueDateMonthplusone = new DateTime(DueDateMonthplusone.Year, DueDateMonthplusone.Month, 1);
                    DateTime DueDateDayminusone = DueDateMonthplusone.AddDays(-1);
                    DueDate = new DateTime(DueDateDayminusone.Year, DueDateDayminusone.Month, DueDateDayminusone.Day);
                }
                //Calculation Method = Days Later for Extension 1
                else
                {
                    DueDate = DueDate.AddDays(valext1DaysLaterDue);
                }
            }
            else if (valTimesExtended.Value == 100000002)  //if (valTimesExtended.Equals(100000002))
            {
                //Calculation Method = Months Later - Specific Day of Month Extension 2
                if (valCalculationMethod.Value == 100000001) //if (valCalculationMethod.Equals(100000001))
                {
                    if (valext2DayDue <= 0)
                    {
                        valext2DayDue = 1;
                    }

                    DateTime DueDateAddMonths = DueDate.AddMonths(valext2MonthsLaterDue);
                    valMonth = DueDateAddMonths.Month;
                    valYear = DueDateAddMonths.Year;

                    calcDate = DueDateAddMonths.Month + "/" + valext2DayDue + "/" + DueDateAddMonths.Year;

                    while (DateTime.TryParse(calcDate, out DueDateAddMonths) == false)
                    {
                        valext2DayDue = valext2DayDue - 1;
                        calcDate = valMonth + "/" + valext2DayDue + "/" + valYear;
                    }

                    DueDate = DueDateAddMonths;

                }
                //Calculation Method = Months Later for Extension 2
                else if (valCalculationMethod.Value == 100000002)
                {
                    DueDate = DueDate.AddMonths(valext2MonthsLaterDue);
                }
                //Calculation Method = Month Later - Last Day of Month Extension 2
                else if (valCalculationMethod.Value == 100000003)
                {
                    DateTime DueDateMonthplusone = DueDate.AddMonths(valext2MonthsLaterDue + 1);
                    DueDateMonthplusone = new DateTime(DueDateMonthplusone.Year, DueDateMonthplusone.Month, 1);
                    DateTime DueDateDayminusone = DueDateMonthplusone.AddDays(-1);
                    DueDate = new DateTime(DueDateDayminusone.Year, DueDateDayminusone.Month, DueDateDayminusone.Day);
                }
                //Calculation Method = Days Later for Extension 2
                else
                {
                    DueDate = DueDate.AddDays(valext2DaysLaterDue);
                }
            }
                result.Set(context, DueDate);
        }
    }
}
