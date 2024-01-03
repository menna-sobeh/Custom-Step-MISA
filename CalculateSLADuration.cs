using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using System.Threading;
using System.Diagnostics;

namespace LinkDev.Crm.Cs.Task
{
    public class CalculateSLADuration : CodeActivity
    {
        #region Global Variables  
        public double WorkingDuration = 0;
        public TimeSpan WorkingStartSpan = new TimeSpan();
        public TimeSpan WorkingEndSpan = new TimeSpan();
        public int Duration = 0;
        public int totalMin = 0;

        #endregion

        #region Input Parameters
        [RequiredArgument]
        [Input("Start Date")]
        public InArgument<DateTime> startDate { get; set; }

        [RequiredArgument]
        [Input("End Date")]
        public InArgument<DateTime> endDate { get; set; }

        [RequiredArgument]
        [Input("SLA")]
        [ReferenceTarget("sla")]
        public InArgument<EntityReference> slaRecord { get; set; }
        #endregion

        #region output Parameter 
        [Output("Calculated Final Duration")]
        public OutArgument<int> FinalDurationInHours { get; set; }

        #endregion
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();          
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);
            ITracingService tracingService = context.GetExtension<ITracingService>();

            tracingService.Trace("custom step started");

            DateTime start = startDate.Get<DateTime>(context); 
            DateTime end = endDate.Get<DateTime>(context);
            EntityReference SLA = slaRecord.Get<EntityReference>(context);

            if(end > start) 
            {
            //Calendar
            tracingService.Trace("1- Get business Calendar Entity");
            Entity slaBusinessHours = service.Retrieve("sla", SLA.Id, new ColumnSet("businesshoursid"));
            EntityReference businessHoursRef = slaBusinessHours.GetAttributeValue<EntityReference>("businesshoursid"); //default calendar lookup

                if (businessHoursRef == null)
                {
                    tracingService.Trace("No Business Hours");
                    TimeSpan timeDiff = end.Subtract(start);

                    if (start.Date == end.Date)//Same Day
                    {
                        totalMin = Convert.ToInt32((timeDiff.TotalMinutes));
                        Duration = totalMin;
                    }
                    else //Diff Days
                    {
                        Duration = Convert.ToInt32((timeDiff.Days * 24 * 60) + (timeDiff.Hours * 60) + (timeDiff.Minutes));
                    }
                }
                else
                {
                    tracingService.Trace("Business Hours Exist");
                    //the schema name of the Business Hours entity is calendar 
                    Entity businessCalendarObject = service.Retrieve("calendar", businessHoursRef.Id, new ColumnSet(true));

                    tracingService.Trace("2- Get Vacations");
                    List<DateTime> vacationList = GetVacations(businessCalendarObject, service, tracingService);

                    tracingService.Trace("3- Working Days Pattern");
                    List<string> workingDaysPatternList = GetWorkingDaysPattern(businessCalendarObject, tracingService);

                    var calendarRules = businessCalendarObject.GetAttributeValue<EntityCollection>("calendarrules");
                    // Get the ID of the inner calendar
                    var innerCalendarId = calendarRules[0].GetAttributeValue<EntityReference>("innercalendarid").Id; //lookup

                    // Retrieve the inner calendar with all of its columns //calendar Entity //12 att
                    var innerCalendar = service.Retrieve("calendar", innerCalendarId, new ColumnSet(true));
                    if (businessCalendarObject != null && businessCalendarObject.Id != Guid.Empty)
                    {
                        // Get the first inner calendar rule 
                        var innerCalendarRule = innerCalendar.GetAttributeValue<EntityCollection>("calendarrules").Entities.FirstOrDefault();

                        var startTime = Convert.ToDouble(innerCalendarRule.GetAttributeValue<int>("offset")) / Convert.ToDouble(60);
                        WorkingDuration = Convert.ToDouble(innerCalendarRule.GetAttributeValue<int>("duration")) / Convert.ToDouble(60);

                        WorkingStartSpan = TimeSpan.FromHours(startTime); //9
                        WorkingEndSpan = TimeSpan.FromHours(startTime + WorkingDuration); //17

                    }
                    tracingService.Trace("4- Get Total Duration");
                    Duration = GetFinalDuration(start, end, workingDaysPatternList, vacationList, tracingService);
                    tracingService.Trace(Duration.ToString());
                    //set field value with the duration
                }

            }
            else
            {
                Duration = 0;
                tracingService.Trace("Duration Can't be negative");
                tracingService.Trace(Duration.ToString());
            }

            FinalDurationInHours.Set(context, Duration);


        }
        public int GetFinalDuration(DateTime startDate, DateTime endDate, List<string> workingDaysPatternList, List<DateTime> vacationList, ITracingService tracingService)
        {
            tracingService.Trace("Inside GetFinalDuration");
            string durationInDays = string.Empty;
            string durationInHours = string.Empty;
            string durationInDaysandHours = string.Empty;
            List<DateTime> originalDatesRangeList = new List<DateTime>();

            #region Prepare Lists 
           
            if ((endDate - startDate).Days < 0)
            {
                endDate = startDate;
                originalDatesRangeList.Add(startDate);
            }
            else
            {
                originalDatesRangeList = Enumerable.Range(0, (endDate - startDate).Days).Select(d => startDate.AddDays(d)).ToList();
                originalDatesRangeList.Add(endDate); //add endDate Seperate to keep the time as it is not as the above list date's time is the same as start date
            }

            tracingService.Trace("1- Exclude Vacation");
            var datesWithoutVacationList = originalDatesRangeList.Where(originalDay => !vacationList.Any(vacationDay => vacationDay.Date == originalDay.Date));

            tracingService.Trace("2- Exclude Weekends");  
            var dateWithoutOffDaysList = datesWithoutVacationList.Where(workingDay => workingDaysPatternList.Any(patternDay => patternDay.ToLower() == workingDay.DayOfWeek.ToString().Substring(0, 2).ToLower()));


            #endregion

            #region Calculate Total Hours
            tracingService.Trace("3- Calculate Total Hours");
            var datesInBetween = dateWithoutOffDaysList.Where(workingDay => workingDay != startDate && workingDay != endDate);

            var totalHoursForDatesInBetween = (datesInBetween.Count() * WorkingDuration);

            TimeSpan totalHoursInStartDate = startDate.Hour > WorkingEndSpan.Hours ? TimeSpan.Zero :
                                    startDate.Hour < WorkingStartSpan.Hours ?
                                    WorkingEndSpan.Subtract(WorkingStartSpan) :
                                    WorkingEndSpan.Subtract(startDate.TimeOfDay);


            TimeSpan totalHoursInEndDate = endDate.Hour < WorkingStartSpan.Hours ?
                                      TimeSpan.Zero :
                                      endDate.Hour > WorkingEndSpan.Hours ?
                                      WorkingEndSpan.Subtract(WorkingStartSpan) :
                                      endDate.TimeOfDay.Subtract(WorkingStartSpan);


            double totalHoursOfStartAndEndDates = 0;

            if (startDate.Date == endDate.Date)
                totalHoursOfStartAndEndDates = (endDate.AddSeconds(-endDate.Second) - startDate.AddSeconds(-startDate.Second)).TotalHours;
            else
                totalHoursOfStartAndEndDates = totalHoursInStartDate.TotalHours + totalHoursInEndDate.TotalHours;
            #endregion

            #region Calculate Days and hours and minutes seperates
            tracingService.Trace("4- Calculate Days and hours and minutes seperates");

            double finalTotalInDays = Convert.ToDouble(datesInBetween.Count());
            double finalTotalHours = 0;
            double finalTotalMinutes = 0;

            if (totalHoursOfStartAndEndDates > WorkingDuration) 
            {

                var totalDiff = totalHoursOfStartAndEndDates / WorkingDuration;

                var daystoAdd = Math.Truncate(totalDiff);

                var hoursAndMins = (totalDiff - daystoAdd) * WorkingDuration;

                finalTotalHours = hoursAndMins > 0 ? Math.Truncate(hoursAndMins) : 0;

                finalTotalMinutes = (hoursAndMins - finalTotalHours) * 60;

                finalTotalInDays += daystoAdd;

            }
            else
            {
                finalTotalHours = Math.Truncate(totalHoursOfStartAndEndDates);

                finalTotalMinutes = (totalHoursOfStartAndEndDates - finalTotalHours) * 60;
            }

            #endregion

            #region Calculate Duration 
            tracingService.Trace("5- Calculate Duration");

            durationInDays = dateWithoutOffDaysList.Count() > 0 ? dateWithoutOffDaysList.Count().ToString() : "1";

            durationInHours = String.Format("{0:0.00}", (totalHoursForDatesInBetween + totalHoursOfStartAndEndDates)).ToString();

            durationInDaysandHours = finalTotalInDays + " d, " + finalTotalHours + " h, " + Math.Truncate(finalTotalMinutes) + " m";
            Duration = Convert.ToInt32((finalTotalInDays * 24 * 60) + (finalTotalHours * 60) + Math.Truncate(finalTotalMinutes));

            #endregion
            tracingService.Trace("6- Finished");

            return Duration;
        }
        
        public  List<DateTime> GetVacations(Entity businessCalendarObject, IOrganizationService OrganizationService, ITracingService tracingService)
        {
            List<DateTime> finalVacationList = new List<DateTime>();

            if (businessCalendarObject != null && businessCalendarObject.Id != Guid.Empty)
            {

                // check that the calendar object has a holiday schedule calendar linked to it
                if (businessCalendarObject.Contains("holidayschedulecalendarid") && businessCalendarObject.Attributes["holidayschedulecalendarid"] != null)
                {

                    var holidaysCalendarEntity = (EntityReference)businessCalendarObject.Attributes["holidayschedulecalendarid"];

                    // retrieve holidays object
                    var holidaysCalendarObject = OrganizationService.Retrieve(holidaysCalendarEntity.LogicalName, holidaysCalendarEntity.Id, new ColumnSet(true));

                    // get calendar rules
                    var calendarRules = holidaysCalendarObject.GetAttributeValue<EntityCollection>("calendarrules");

                    foreach (var calenderRule in calendarRules.Entities)
                    {
                        List<DateTime> vacationList = new List<DateTime>();

                        var startDate = (DateTime)calenderRule.Attributes["effectiveintervalstart"];
                        var endDate = (DateTime)calenderRule.Attributes["effectiveintervalend"];

                        vacationList = Enumerable.Range(0, (endDate - startDate).Days).Select(d => startDate.AddDays(d)).ToList();

                        finalVacationList.AddRange(vacationList);
                    }
                }

            }


            return finalVacationList.OrderBy(x => x.Date).ToList();
        }
        public  List<string> GetWorkingDaysPattern(Entity businessCalendarObject, ITracingService tracingService)
        {

            List<string> patternDaysShortName = new List<string>();

            if (businessCalendarObject != null && businessCalendarObject.Id != Guid.Empty)
            {
                var calendarRules = businessCalendarObject.GetAttributeValue<EntityCollection>("calendarrules");

                var firstRulePattern = calendarRules[0].GetAttributeValue<string>("pattern");
                // FREQ=WEEKLY;INTERVAL=1;BYDAY=SU,MO,TU,WE,TH,FR

                patternDaysShortName = firstRulePattern.Substring(firstRulePattern.LastIndexOf("BYDAY")).Split(new string[] { "BYDAY=", "," }, StringSplitOptions.RemoveEmptyEntries).ToList();

            }
            return patternDaysShortName;
        }

    }
}
