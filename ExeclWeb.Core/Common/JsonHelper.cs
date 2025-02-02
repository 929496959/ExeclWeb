﻿using Newtonsoft.Json;

namespace ExeclWeb.Core.Common
{
    public static class JsonHelper
    {
        public static string ToJson(this object obj)
        {
            var jsonSetting = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                Formatting = Formatting.None
            };
            return obj == null ? default : JsonConvert.SerializeObject(obj, jsonSetting);
        }
        public static T ToObject<T>(this string json)
        {
            return json == null ? default : JsonConvert.DeserializeObject<T>(json);
        }
    }
}