using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UbikeInfoCatcher.Model
{
    public class CityBikeStation
    {
        public string StationID { get; set; }
        public string StationNO { get; set; }
        public string StationPic { get; set; }
        public string StationPic2 { get; set; }
        public string StationPic3 { get; set; }
        public string StationMap { get; set; }
        public string StationName { get; set; }
        public string StationAddress { get; set; }
        public string StationLat { get; set; }
        public string StationLon { get; set; }
        public string StationDesc { get; set; }
        public string StationNums1 { get; set; } //目前數量
        public string StationNums2 { get; set; } //尚餘空位
    }
}
