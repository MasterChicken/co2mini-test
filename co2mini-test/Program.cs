using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.Console;
using HidLibrary;

namespace Co2miniTest
{
    public static class Program
    {
        /// <summary>ベンダーID</summary>
        private static readonly int CO2MINI_VENDOR_ID = 0x04d9;
        /// <summary>プロダクトID</summary>
        private static readonly int CO2MINI_PRODUCT_ID = 0xa052;
        /// <summary>初期化に使用するマジックナンバー</summary>
        private static readonly byte[] INIT_MAGIC_NUMBERS = new byte[] { 0x00, 0xC4, 0xC6, 0xC0, 0x92, 0x40, 0x23, 0xDC, 0x96 };
        /// <summary>解読時に使用するマジックナンバー1</summary>
        private static readonly int[] DECRYPT_MAGIC_NUMBERS_1 = new int[] { 0xC4, 0xC6, 0xC0, 0x92, 0x40, 0x23, 0xDC, 0x96 };
        /// <summary>解読時に使用するマジックナンバー2</summary>
        private static readonly int[] DECRYPT_MAGIC_NUMBERS_2 = new int[] { 0x48, 0x74, 0x65, 0x6D, 0x70, 0x39, 0x39, 0x65 };
        /// <summary>解読時に使用するマジックナンバー3</summary>
        private static readonly int[] DECRYPT_MAGIC_NUMBERS_3 = new int[] { 0x84, 0x47, 0x56, 0xD6, 0x07, 0x93, 0x93, 0x56 };
        /// <summary>解読時に使用するデータのシャッフル順</summary>
        private static readonly int[] DATA_SHUFFLE_ORDER = new int[] { 2, 4, 0, 7, 1, 6, 5, 3 };
        /// <summary>データの既定の長さ</summary>
        private static readonly int DATA_LENGTH = 8;
        /// <summary>インバリアントカルチャ</summary>
        private static readonly System.Globalization.CultureInfo INVALIANT_CULTURE = System.Globalization.CultureInfo.InvariantCulture;
        /// <summary>データが揃うのを待つ時間</summary>
        private static readonly TimeSpan DATA_TIMEOUT = TimeSpan.FromSeconds(10);

        private enum ReadDataType
        {
            None,
            Co2,
            Temperature
        }

        /// <summary>
        /// 処理対象のデバイス
        /// </summary>
        public static HidDevice device = null;

        /// <summary>
        /// データの収集完了を待つシグナル
        /// </summary>
        public static ManualResetEventSlim dataCollectedSignal = null;

        /// <summary>
        /// すでに終了処理が始まっているかどうか
        /// </summary>
        public static bool quitting = false;

        /// <summary>
        /// CO2濃度（PPM）
        /// </summary>
        public static int? co2 = null;

        /// <summary>
        /// 気温（セルシウス度）
        /// </summary>
        public static double? temperature = null;

        public static void Main()
        {
            // 例外処理定義
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            try
            {
                // デバイスを検索
                var devices = HidDevices.Enumerate(CO2MINI_VENDOR_ID, new int[] { CO2MINI_PRODUCT_ID });

                if (!devices.Any())
                {
                    throw new Exception("no co2mini device found");
                }
                if (devices.Count() != 1)
                {
                    throw new Exception("multiple co2mini device found");
                }

                device = devices.First();
                device.MonitorDeviceEvents = true;
                device.OpenDevice();

                 // データ収集待ちシグナル
                dataCollectedSignal = new ManualResetEventSlim();

                if (!device.IsOpen || !device.IsConnected)
                {
                    throw new Exception("device is neither opened nor connected");
                }

                // 初期化要求送信
                if (!device.WriteFeatureData(INIT_MAGIC_NUMBERS))
                {
                    throw new Exception("failed to write feature report while initialization");
                }

                // データ読み取り
                device.ReadReport(DeviceOnReport, Convert.ToInt32(DATA_TIMEOUT.TotalMilliseconds));
                dataCollectedSignal.Wait();

                // 結果表示
                var resultData = new ResultData
                {
                    IsOK = true,
                    Co2 = co2,
                    Temperature = temperature
                };
                Write(resultData.ToJson());

            }
            finally
            {
                quitting = true;

                // シグナルを止める
                if (dataCollectedSignal != null)
                {
                    dataCollectedSignal.Dispose();
                }
                dataCollectedSignal = null;

                // デバイスを閉じる
                if (device != null)
                {
                    device.CloseDevice();
                    device.Dispose();
                }
                device = null;
            }


#if DEBUG
            ReadLine();
#endif

            return;
        }

        /// <summary>
        /// デバイスから送られたレポートを解析します。
        /// </summary>
        /// <param name="report"></param>
        public static void DeviceOnReport(HidReport report)
        {
            // 終了処理中
            if (quitting)
            {
                return;
            }

            // デバイス存在確認
            if (device == null || !device.IsOpen || !device.IsConnected)
            {
                goto fail;
            }

            // レポート存在確認
            if (report == null || !report.Exists)
            {
                goto fail;
            }

            // レポート読み取り
            var rawData = report.Data;
            if (rawData == null || rawData.Length != DATA_LENGTH)
            {
                goto fail;
            }

            try
            {
                var decryptedData = DecryptData(rawData);
                var (type, value) = ReadData(decryptedData);

                switch (type)
                {
                    case ReadDataType.Co2:
                        co2 = Convert.ToInt32(value);
                        break;
                    case ReadDataType.Temperature:
                        temperature = Math.Round(value * 10000) / 10000;
                        break;
                    default:
                        throw new NotImplementedException();
                }

            }
            // 無視するエラー
            catch (UnsupportedDataException) { }
            catch (DataDecryptionException) { }

        fail:
            if (co2.HasValue && temperature.HasValue)
            {
                // データ収集完了
                dataCollectedSignal.Set();
            }
            else if (device != null && device.IsOpen && device.IsConnected)
            {
                // 未完了につき、まだ待つ
                device.ReadReport(DeviceOnReport, Convert.ToInt32(DATA_TIMEOUT.TotalMilliseconds));
            }
        }

        /// <summary>
        /// 受信したデータの暗号化を解除し、破損していないデータを返します。
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        static int[] DecryptData(byte[] data)
        {
            var p1 = new int[DATA_LENGTH];
            var p2 = new int[DATA_LENGTH];
            var p3 = new int[DATA_LENGTH];
            var p4 = new int[DATA_LENGTH];

            // 並び替え
            for (var i = 0; i < DATA_LENGTH; ++i)
            {
                p1[DATA_SHUFFLE_ORDER[i]] = data[i];
            }

            // XOR
            for (var i = 0; i < DATA_LENGTH; ++i)
            {
                p2[i] = p1[i] ^ DECRYPT_MAGIC_NUMBERS_1[i];
            }

            // ビットシフト
            for (var i = 0; i < DATA_LENGTH; ++i)
            {
                p3[i] = ((p2[i] >> 3) | (p2[(i - 1 + 8) % 8] << 5)) & 0xff;
            }

            // 差し引き
            for (var i = 0; i < DATA_LENGTH; ++i)
            {
                p4[i] = ((0x100 + p3[i] - DECRYPT_MAGIC_NUMBERS_3[i]) & 0xff);
            }

            // チェックサム
            if (p4[4] != 0x0d || ((p4[0] + p4[1] + p4[2]) & 0xff) != p4[3])
            {
                throw new DataDecryptionException("decrypted data has invalid checksum");
            }

            return p4;
        }

        /// <summary>
        /// データに含まれている値を読み取ります。
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        static (ReadDataType type, double value) ReadData(int[] data)
        {
            var op = data[0];
            var value = data[1] << 8 | data[2];

            if (op == 0x50)
            {
                // co2
                return (ReadDataType.Co2, value);
            }
            else if (op == 0x42)
            {
                // temperature
                return (ReadDataType.Temperature, value / 16.0 - 273.15);
            }
            else
            {
                // サポートされていないデータ
                throw new UnsupportedDataException();
            }
        }



        /// <summary>
        /// 例外エラー処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            quitting = true;

            var resultData = new ResultData
            {
                IsOK = false,
            };
            if (e.ExceptionObject != null)
            {
                resultData.Message = e.ExceptionObject.ToString();
            }
            else
            {
                resultData.Message = "unknown error occured";
            }
            Write(resultData.ToJson());

            // シグナルを止める
            if (dataCollectedSignal != null)
            {
                dataCollectedSignal.Dispose();
            }
            dataCollectedSignal = null;

            // デバイスを閉じる
            if (device != null)
            {
                device.CloseDevice();
                device.Dispose();
            }
            device = null;

            // エラー終了
            Environment.Exit(1);
        }
    }

    [Serializable]
    public class DataDecryptionException : Exception
    {
        public DataDecryptionException() { }
        public DataDecryptionException(string message) : base(message) { }
        public DataDecryptionException(string message, Exception inner) : base(message, inner) { }
        protected DataDecryptionException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }


    [Serializable]
    public class UnsupportedDataException : Exception
    {
        public UnsupportedDataException() { }
        public UnsupportedDataException(string message) : base(message) { }
        public UnsupportedDataException(string message, Exception inner) : base(message, inner) { }
        protected UnsupportedDataException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
