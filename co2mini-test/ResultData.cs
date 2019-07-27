using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.IO;

namespace Co2miniTest
{
    /// <summary>
    /// 取得結果をセットするオブジェクト
    /// </summary>
    [DataContract]
    public class ResultData
    {
        /// <summary>
        /// 取得に成功したかどうか
        /// </summary>
        [DataMember(EmitDefaultValue = true, IsRequired = true, Name = "isOk")]
        public bool IsOK { get; set; } = false;

        /// <summary>
        /// 取得に際して発生したメッセージ
        /// </summary>
        [DataMember(EmitDefaultValue = true, Name = "message")]
        public string Message { get; set; } = null;

        /// <summary>
        /// 二酸化炭素濃度（PPM）
        /// </summary>
        [DataMember(EmitDefaultValue = false, Name = "co2")]
        public int? Co2 { get; set; } = null;

        /// <summary>
        /// 気温（セルシウス度）
        /// </summary>
        [DataMember(EmitDefaultValue = false, Name = "temperature")]
        public double? Temperature { get; set; } = null;

        /// <summary>
        /// 現在のオブジェクト意をJSON文字列にして返します。
        /// </summary>
        /// <returns></returns>
        public string ToJson()
        {
            using (var memoryStream = new MemoryStream())
            using (var streamReader = new StreamReader(memoryStream))
            {
                var serializer = new DataContractJsonSerializer(typeof(ResultData));
                serializer.WriteObject(memoryStream, this);
                memoryStream.Position = 0;

                return streamReader.ReadToEnd();
            }
        }
    }
}
