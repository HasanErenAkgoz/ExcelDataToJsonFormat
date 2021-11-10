using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataToJsonFormat
{
   public class JsonClass
    {
        [JsonProperty("label")]
        public string label { get; set; }

        [JsonProperty("tr-TR")]
        public string TrTR { get; set; }

        [JsonProperty("en-US")]
        public string EnUS { get; set; }

        [JsonProperty("fr-FR")]
        public string FrFR { get; set; }

        [JsonProperty("it-IT")]
        public string ItIT { get; set; }

        [JsonProperty("bg-BG")]
        public string BgBG { get; set; }

        [JsonProperty("ro-RO")]
        public string RoRO { get; set; }

        [JsonProperty("ru-RU")]
        public string RuRU { get; set; }
    }
}
