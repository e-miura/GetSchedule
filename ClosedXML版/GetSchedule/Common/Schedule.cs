using System;
namespace GetSchedule.Common
{
    struct Schedule
    {
        public string date;
        public string jobid;
        public string body;
        public TimeSpan time;

        public Schedule(string date, string jobid, string body, TimeSpan time)
        {
            this.date = date;
            this.jobid = jobid;
            this.body = body;
            this.time = time;
        }
    }

    struct SortSchedule
    {
        public string date;
        public string jobid;
        public string body;
        public TimeSpan time;

        public SortSchedule(string date, string jobid, string body, TimeSpan time)
        {
            this.date = date;
            this.jobid = jobid;
            this.body = body;
            this.time = time;
        }
    }
}
