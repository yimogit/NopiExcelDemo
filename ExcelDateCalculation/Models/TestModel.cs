using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDateCalculation.Models
{
    public class TestModel
    {
        public string No { get; set; }
        public DateTime BeginTime { get; set; }
        public DateTime EndTime { get; set; }
        public string Result { get; set; }
    }
    public class TestModelExt
    {
        public TimeSpan GetDifferTime(DateTime _begin, DateTime _endTime)
        {
            var t1 = new DateTime(_endTime.Year, _endTime.Month, _endTime.Day, 9, 0, 0);
            var t2 = new DateTime(_endTime.Year, _endTime.Month, _endTime.Day, 12, 30, 0);
            var t3 = new DateTime(_endTime.Year, _endTime.Month, _endTime.Day, 13, 30, 0);
            var t4 = new DateTime(_endTime.Year, _endTime.Month, _endTime.Day, 18, 0, 0);


            if (_begin > _endTime)//结束时间大于开始时间->隔了一天
            {
                var ts = new TimeSpan();
                if (_begin < t4)
                {
                    ts = GetDifferTime(_begin, t4);
                }
                if (t1 < _endTime)
                {
                    //头一天的时间
                    ts = ts.Add(GetDifferTime(t1, _endTime));
                }
                return ts;
            }
            if (_begin < t1)
            {
                _begin = t1;
            }
            if (_begin > t2 && _begin < t3)
            {
                _begin = t3;
            }
            if (_begin > t4)
            {
                _begin = t4;
            }
            if (_endTime < t1)
            {
                _endTime = t1;
            }
            if (_endTime > t2 && _endTime < t3)
            {
                _endTime = t2;
            }
            if (_endTime > t4)
            {
                _endTime = t4;
            }
            if (_begin < _endTime)
            {
                return _endTime - _begin;
            }
            return new TimeSpan();

        }
        public string GetJishuanResult(TestModel model)
        {
            var _beginTime = model.BeginTime;
            var _endTime = model.EndTime;
            var toastTime = new TimeSpan();
            if (_beginTime > _endTime)
            {
                return "数据有问题的节奏啊";
            }
            var diffDay = (_endTime - _beginTime).Days;

            if (diffDay == 0)
            {
                toastTime = GetDifferTime(_beginTime, _endTime);
            }
            else
            {
                var ts8 = new TimeSpan(8, 0, 0);
                for (int i = 0; i < diffDay; i++)
                {
                    if (i + 1 == diffDay)
                    {
                        var daySpan = GetDifferTime(_beginTime.AddDays(i), _endTime);
                        toastTime = toastTime.Add(daySpan);
                    }
                    else
                    {
                        toastTime = toastTime.Add(ts8);
                    }
                }
            }
            return (toastTime.Days * 24 + toastTime.Hours) + ":" + toastTime.Minutes + ":" + toastTime.Seconds;
        }
    }
}
