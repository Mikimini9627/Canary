using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Canary
{
    public class Parameter1
    {
        public string[] prefectureIds { get; set; }
        public object[] cityIds { get; set; }
        public object[] chomeiIds { get; set; }
        public object[] stationIds { get; set; }
        public Rentmin rentMin { get; set; }
        public Rentmax rentMax { get; set; }
        public bool includeAdminFee { get; set; }
        public Squaremin squareMin { get; set; }
        public Squaremax squareMax { get; set; }
        public Oldmax oldMax { get; set; }
        public Duringmax duringMax { get; set; }
        public string[] searchOptionIds { get; set; }
        public object[] keywords { get; set; }
        public string[] layoutNames { get; set; }
        public bool isNewArrival { get; set; }
        public Commute[] commutes { get; set; }
        public Shatakucode shatakuCode { get; set; }
    }

    public class Parameter2
    {
        public string[] prefectureIds { get; set; }
        public object[] cityIds { get; set; }
        public object[] chomeiIds { get; set; }
        public object[] stationIds { get; set; }
        public Rentmin rentMin { get; set; }
        public Rentmax rentMax { get; set; }
        public bool includeAdminFee { get; set; }
        public Squaremin squareMin { get; set; }
        public Squaremax squareMax { get; set; }
        public Oldmax oldMax { get; set; }
        public Duringmax duringMax { get; set; }
        public string[] searchOptionIds { get; set; }
        public object[] keywords { get; set; }
        public string[] layoutNames { get; set; }
        public bool isNewArrival { get; set; }
        public Commute[] commutes { get; set; }
        public Shatakucode shatakuCode { get; set; }

        public int limit { get; set; }

        public int listType { get; set; }

        public Offset offset { get; set; }

        public int sortType { get; set; }

        public bool searchPurpose { get; set; }

        public string searchSessionId { get; set; }

        public Parameter2()
        {

        }

        public Parameter2(Parameter2 source)
        {
            prefectureIds = new string[source.prefectureIds.Length];
            for(int i = 0; i < source.prefectureIds.Length; i++)
            {
                prefectureIds[i] = source.prefectureIds[i];
            }

            cityIds = new object[source.cityIds.Length];
            for (int i = 0; i < source.cityIds.Length; i++)
            {
                cityIds[i] = source.cityIds[i];
            }

            chomeiIds = new object[source.chomeiIds.Length];
            for (int i = 0; i < source.chomeiIds.Length; i++)
            {
                chomeiIds[i] = source.chomeiIds[i];
            }

            stationIds = new object[source.stationIds.Length];
            for (int i = 0; i < source.stationIds.Length; i++)
            {
                stationIds[i] = source.stationIds[i];
            }

            keywords = new object[source.keywords.Length];
            for (int i = 0; i < source.keywords.Length; i++)
            {
                keywords[i] = source.keywords[i];
            }

            searchOptionIds = new string[source.searchOptionIds.Length];
            for (int i = 0; i < source.searchOptionIds.Length; i++)
            {
                searchOptionIds[i] = source.searchOptionIds[i];
            }

            layoutNames = new string[source.layoutNames.Length];
            for (int i = 0; i < source.layoutNames.Length; i++)
            {
                layoutNames[i] = source.layoutNames[i];
            }

            rentMin = new Rentmin(source.rentMin);
            rentMax = new Rentmax(source.rentMax);
            squareMin = new Squaremin(source.squareMin);
            squareMax = new Squaremax(source.squareMax);
            oldMax = new Oldmax(source.oldMax);
            duringMax = new Duringmax(source.duringMax);
            commutes = new Commute[source.commutes.Length];
            for (int i = 0; i < source.commutes.Length; i++)
            {
                commutes[i] = new Commute(source.commutes[i]);
            }
            shatakuCode = new Shatakucode(source.shatakuCode);
            offset = new Offset(source.offset);

            searchSessionId = source.searchSessionId;
            limit = source.limit;
            sortType = source.sortType;
            listType = source.listType;
            searchPurpose = source.searchPurpose;
            isNewArrival = source.isNewArrival;
            includeAdminFee = source.includeAdminFee;
        }
    }

    public class Offset
    {
        public string value { get; set; }
        public bool hasValue { get; set; }

        public Offset()
        {

        }

        public Offset(Offset source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class Rentmin
    {
        public int value { get; set; }
        public bool hasValue { get; set; }

        public Rentmin()
        {

        }

        public Rentmin(Rentmin source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class Rentmax
    {
        public int value { get; set; }
        public bool hasValue { get; set; }

        public Rentmax()
        {

        }

        public Rentmax(Rentmax source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class Squaremin
    {
        public int value { get; set; }
        public bool hasValue { get; set; }

        public Squaremin()
        {

        }

        public Squaremin(Squaremin source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class Squaremax
    {
        public int value { get; set; }
        public bool hasValue { get; set; }

        public Squaremax()
        {

        }

        public Squaremax(Squaremax source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class Oldmax
    {
        public int value { get; set; }
        public bool hasValue { get; set; }

        public Oldmax()
        {

        }

        public Oldmax(Oldmax source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class Duringmax
    {
        public int value { get; set; }
        public bool hasValue { get; set; }

        public Duringmax()
        {

        }

        public Duringmax(Duringmax source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class Shatakucode
    {
        public string value { get; set; }
        public bool hasValue { get; set; }

        public Shatakucode()
        {

        }

        public Shatakucode(Shatakucode source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class Commute
    {
        public string sourceStationId { get; set; }
        public int timeMinutes { get; set; }
        public Changecount changeCount { get; set; }

        public Commute()
        {

        }

        public Commute(Commute source)
        {
            sourceStationId = source.sourceStationId;
            timeMinutes = source.timeMinutes;
            changeCount = new Changecount(source.changeCount);
        }
    }

    public class Changecount
    {
        public int value { get; set; }
        public bool hasValue { get; set; }

        public Changecount()
        {

        }

        public Changecount(Changecount source)
        {
            value = source.value;
            hasValue = source.hasValue;
        }
    }

    public class TotalCount
    {
        public int totalCount { get; set; }
    }

}
