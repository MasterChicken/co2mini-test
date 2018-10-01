using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using HidLibrary;

namespace co2mini_test
{
    public class DataDecryptionException : Exception
    {
        public DataDecryptionException() : base() { }
        public DataDecryptionException(string message) : base(message) { }
    }


    public static class Program
    {
        /// <summary>ベンダーID</summary>
        private static readonly int CO2MINI_VENDOR_ID = 0x04d9;
        /// <summary>プロダクトID</summary>
        private static readonly int CO2MINI_PRODUCT_ID = 0xa052;
        /// <summary>初期化＆解読時に使用するマジックナンバー</summary>
        private static readonly byte[] MAGIC_NUMBERS_1 = new byte[] { 0xC4, 0xC6, 0xC0, 0x92, 0x40, 0x23, 0xDC, 0x96 };
        /// <summary>解読時に使用するマジックナンバー2</summary>
        private static readonly int[] MAGIC_NUMBERS_2 = new int[] { 0x48, 0x74, 0x65, 0x6D, 0x70, 0x39, 0x39, 0x65 };
        /// <summary>解読時に使用するマジックナンバー3</summary>
        private static readonly int[] MAGIC_NUMBERS_3 = new int[] { 0x84, 0x47, 0x56, 0xD6, 0x07, 0x93, 0x93, 0x56 };
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
            CO2,
            Temperature
        }


        /// <summary>
        /// 受信したデータの暗号化を解除し、破損していないデータを返します。
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        static int[] DecryptData(byte[] data)
        {
            if (data == null) { throw new ArgumentNullException(nameof(data)); }
            if (data.Length != DATA_LENGTH) { throw new ArgumentException($"{nameof(data)}.Length != {DATA_LENGTH}"); }

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
                p2[i] = p1[i] ^ MAGIC_NUMBERS_1[i];
            }

            // ビットシフト
            for (var i = 0; i < DATA_LENGTH; ++i)
            {
                p3[i] = ((p2[i] >> 3) | (p2[(i - 1 + 8) % 8] << 5)) & 0xff;
            }

            // 差し引き
            for (var i = 0; i < DATA_LENGTH; ++i)
            {
                p4[i] = ((0x100 + p3[i] - MAGIC_NUMBERS_3[i]) & 0xff);
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
            if (data == null) { throw new ArgumentNullException(nameof(data)); }
            if (data.Length != DATA_LENGTH) { throw new ArgumentException($"{nameof(data)}.Length != {DATA_LENGTH}"); }

            var op = data[0];
            var value = data[1] << 8 | data[2];

            if (op == 0x50)
            {
                // co2
                return (ReadDataType.CO2, value);
            }
            else if (op == 0x42)
            {
                // temperature
                return (ReadDataType.Temperature, value / 16.0 - 273.15);
            }
            else
            {
                // none
                return (ReadDataType.None, 0);
            }
        }

        static int Main()
        {
            try
            {
                // デバイスを検索
                var devices = HidDevices.Enumerate(CO2MINI_VENDOR_ID, new int[] { CO2MINI_PRODUCT_ID });

                if (!devices.Any()) { throw new Exception("no valid device found"); }
                if (devices.Count() != 1) { throw new Exception("multiple devices currently not supported"); }

                // 今のところ最初の1台にだけ対応
                var device = devices.First();
                int? valueCO2 = null;
                double? valueTemp = null;

                CountdownEvent waitDataSignal = null;
                System.Timers.Timer timer = null;
                DateTime startTime = DateTime.MinValue;

                device.MonitorDeviceEvents = true;
                device.OpenDevice();
                try
                {
                    if (!device.WriteFeatureData(MAGIC_NUMBERS_1.Prepend((byte)0x0).ToArray()))
                    {
                        throw new Exception("writing feature report failed");
                    }

                    // データが来た際のコールバック
                    void OnReport(HidReport report)
                    {
                        if (!device.IsOpen || !device.IsConnected || waitDataSignal.IsSet) { return; }

                        if (report.Exists)
                        {
                            var rawBytes = report.Data;
                            if (rawBytes != null && rawBytes.Length == DATA_LENGTH)
                            {
                                try
                                {
                                    var rawData = DecryptData(rawBytes);
                                    var (type, value) = ReadData(rawData);

                                    if (type == ReadDataType.CO2)
                                    {
                                        valueCO2 = (int)value;
                                    }
                                    else if (type == ReadDataType.Temperature)
                                    {
                                        valueTemp = value;
                                    }
                                }
                                catch (DataDecryptionException) { }
                            }
                        }

                        if (valueCO2.HasValue && valueTemp.HasValue)
                        {
                            // データが揃った
                            waitDataSignal.Signal();
                        }
                        else
                        {
                            // まだ待つ
                            device.ReadReport(OnReport);
                        }
                    }

                    waitDataSignal = new CountdownEvent(1);

                    timer = new System.Timers.Timer()
                    {
                        Enabled = false,
                        AutoReset = false,
                        Interval = 1000,
                    };
                    timer.Elapsed += (sender, e) =>
                    {
                        timer.Enabled = false;
                        if (waitDataSignal.IsSet) { return; }

                        if ((e.SignalTime - startTime) > DATA_TIMEOUT)
                        {
                            // 時間切れ
                            waitDataSignal.Signal();
                        }
                        else
                        {
                            // まだ待つ
                            timer.Enabled = true;
                        }
                    };

                    startTime = DateTime.Now;
                    timer.Enabled = true;
                    device.ReadReport(OnReport);

                    waitDataSignal.Wait();
                    timer.Enabled = false;


                }
                finally
                {
                    device.CloseDevice();
                }

                // 結果を書き出す
                Console.Write("co2:");
                if (valueCO2.HasValue)
                {
                    Console.Write(valueCO2.Value.ToString("0", INVALIANT_CULTURE));
                }
                else
                {
                    Console.Write("null");
                }
                Console.Write("\t");

                Console.Write("temp:");
                if (valueTemp.HasValue)
                {
                    Console.Write(valueTemp.Value.ToString("0.0#####", INVALIANT_CULTURE));
                }
                else
                {
                    Console.Write("null");
                }
                Console.Write("\t");
                Console.WriteLine();

#if DEBUG
                Console.ReadKey();
#endif

                return 0;

            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.ToString());
                return 1;
            }

        }
    }
}
