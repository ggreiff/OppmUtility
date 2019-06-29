using System;
using System.Collections.Generic;
using System.Linq;

namespace OppmUtility.Utility
{
    public class SubKeyMap
    {
        public String Category { get; set; }

        public String KeyValue { get; set; }

        public String CellValue { get; set; }

        public Boolean IsMatch()
        {
            return KeyValue.IsTrimEqualTo(CellValue, true);
        }
    }

    public class MapList
    {
        public List<SubKeyMap> KeyMaps { get; set; }

        public MapList()
        {
            KeyMaps = new List<SubKeyMap>();
        }

        public Boolean CheckMap
        {
            get
            {
                return KeyMaps.HasItems() && KeyMaps.All(subKeyMap => subKeyMap.IsMatch());
            }
        }
    }
}
