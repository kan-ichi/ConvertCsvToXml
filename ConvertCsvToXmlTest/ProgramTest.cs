using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ConvertCsvToXmlTest
{
    [TestClass]
    public class ProgramTest
    {
        /// <summary>
        /// メイン処理のテスト
        /// ・入力ファイルの拡張子が「csv」でない場合、読み飛ばされる
        /// ・入力CSVと出力XMLの内容が一致する
        /// </summary>
        [TestMethod]
        public void TestMethod1()
        {
            string testingFolderName = Path.Combine(GetAppPath(), "TestData");

            string inputCsvPathX;
            {
                string inputCsvFileName = MethodBase.GetCurrentMethod().Name + ".csvX";
                inputCsvPathX = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] { };
                File.WriteAllLines(inputCsvPathX, lines);
            }

            string inputCsvPath;
            {
                string inputCsvFileName = MethodBase.GetCurrentMethod().Name + ".CSV";
                inputCsvPath = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] {
                    "1a,\"1,b\",1c",
                    "2a,,2c,2d",
                    "3a, "
                };
                File.WriteAllLines(inputCsvPath, lines);
            }

            string outputXmlPath;
            {
                string outputXmlFileName = MethodBase.GetCurrentMethod().Name + ".xml";
                outputXmlPath = Path.Combine(testingFolderName, outputXmlFileName);
            }

            ConvertCsvToXml.Program.Request request = new ConvertCsvToXml.Program.Request();
            {
                request.InputCsvFolderName = testingFolderName;
                request.InputCsvEncoding = "Shift_JIS";
                request.XsdFilePath = Path.Combine(testingFolderName, "ConvertCsvToXml.exe.xsd");
                request.OutputXmlFileNameByArgs = outputXmlPath;
                request.OutputXmlEncoding = "UTF-8";
            }

            {
                var target = new PrivateType(typeof(ConvertCsvToXml.Program));
                target.InvokeStatic("MainProcess", request);
            }

            {
                DataSet ds = new DataSet();
                ds.ReadXml(outputXmlPath);
                DataTable outputXmlTable = ds.Tables[0];

                Assert.AreEqual(3, outputXmlTable.Rows.Count);
                Assert.AreEqual("1a", Convert.ToString(outputXmlTable.Rows[0][0]));
                Assert.AreEqual("1,b", Convert.ToString(outputXmlTable.Rows[0][1]));
                Assert.AreEqual("1c", Convert.ToString(outputXmlTable.Rows[0][2]));
                Assert.AreEqual("2a", Convert.ToString(outputXmlTable.Rows[1][0]));
                Assert.AreEqual("", Convert.ToString(outputXmlTable.Rows[1][1]));
                Assert.AreEqual("2c", Convert.ToString(outputXmlTable.Rows[1][2]));
                Assert.AreEqual("3a", Convert.ToString(outputXmlTable.Rows[2][0]));
                Assert.AreEqual(" ", Convert.ToString(outputXmlTable.Rows[2][1]));
                Assert.AreEqual("", Convert.ToString(outputXmlTable.Rows[2][2]));
            }

            File.Delete(inputCsvPathX);
            File.Delete(inputCsvPath);
            File.Delete(outputXmlPath);
        }

        /// <summary>
        /// エントリーポイントからメイン処理のテスト
        /// ・入力CSVのレコードがない場合、出力XMLのレコードの行要素数もゼロとなる
        /// ・入力CSVのファイル名が指定されていない場合、入力CSVのフォルダのすべてのファイルが処理対象となる
        /// </summary>
        [TestMethod]
        public void TestMethod2()
        {
            string exeFolderName = GetAppPath();
            string testingFolderName = Path.Combine(GetAppPath(), "TestData");

            string inputCsvPath;
            {
                string inputCsvFileName = MethodBase.GetCurrentMethod().Name + ".csv";
                inputCsvPath = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] { };
                File.WriteAllLines(inputCsvPath, lines);
            }

            string outputXmlPath;
            {
                string outputXmlFileName = MethodBase.GetCurrentMethod().Name + ".xml";
                outputXmlPath = Path.Combine(testingFolderName, outputXmlFileName);
            }

            {
                File.Copy(Path.Combine(testingFolderName, "ConvertCsvToXml.exe.xsd"), Path.Combine(exeFolderName, "ConvertCsvToXml.exe.xsd"), true);
                string[] args = new string[] { inputCsvPath, outputXmlPath };
                ConvertCsvToXml.Program.Main(args);
                File.Delete(Path.Combine(exeFolderName, "ConvertCsvToXml.exe.xsd"));
            }

            {
                DataSet ds = new DataSet();
                ds.ReadXml(outputXmlPath);

                Assert.AreEqual(0, ds.Tables.Count);
            }

            File.Delete(inputCsvPath);
            File.Delete(outputXmlPath);
        }

        /// <summary>
        /// エントリーポイントのテスト
        /// ・入力ファイル名がブランクであるため、例外発生
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(System.ArgumentException))]
        public void TestMethod3()
        {
            string[] args = new string[] { };
            ConvertCsvToXml.Program.Main(args);
        }

        /// <summary>
        /// メイン処理のテスト
        /// ・入力CSVの最初の一行を読込対象外とした場合、最初の一行が読み飛ばされる
        /// </summary>
        [TestMethod]
        public void TestMethod4()
        {
            string testingFolderName = Path.Combine(GetAppPath(), "TestData");

            string inputCsvPath;
            {
                string inputCsvFileName = MethodBase.GetCurrentMethod().Name + ".CSV";
                inputCsvPath = Path.Combine(testingFolderName, inputCsvFileName);
                string[] lines = new string[] {
                    "1a,\"1,b\",1c",
                    "2a,,2c,2d",
                    "3a, "
                };
                File.WriteAllLines(inputCsvPath, lines);
            }

            string outputXmlPath;
            {
                string outputXmlFileName = MethodBase.GetCurrentMethod().Name + ".xml";
                outputXmlPath = Path.Combine(testingFolderName, outputXmlFileName);
            }

            ConvertCsvToXml.Program.Request request = new ConvertCsvToXml.Program.Request();
            {
                request.InputCsvFolderName = testingFolderName;
                request.InputCsvEncoding = "Shift_JIS";
                request.InputCsvSkipFirstRow = true;
                request.XsdFilePath = Path.Combine(testingFolderName, "ConvertCsvToXml.exe.xsd");
                request.OutputXmlFileNameByArgs = outputXmlPath;
                request.OutputXmlEncoding = "UTF-8";
            }

            {
                var target = new PrivateType(typeof(ConvertCsvToXml.Program));
                target.InvokeStatic("MainProcess", request);
            }

            {
                DataSet ds = new DataSet();
                ds.ReadXml(outputXmlPath);
                DataTable outputXmlTable = ds.Tables[0];

                Assert.AreEqual(2, outputXmlTable.Rows.Count);
                Assert.AreEqual("2a", Convert.ToString(outputXmlTable.Rows[0][0]));
                Assert.AreEqual("", Convert.ToString(outputXmlTable.Rows[0][1]));
                Assert.AreEqual("2c", Convert.ToString(outputXmlTable.Rows[0][2]));
                Assert.AreEqual("3a", Convert.ToString(outputXmlTable.Rows[1][0]));
                Assert.AreEqual(" ", Convert.ToString(outputXmlTable.Rows[1][1]));
                Assert.AreEqual("", Convert.ToString(outputXmlTable.Rows[1][2]));
            }

            File.Delete(inputCsvPath);
            File.Delete(outputXmlPath);

        }

        /// <summary>
        /// テスト実施中のフォルダパスを取得します
        /// </summary>
        private static string GetAppPath()
        {
            string path = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase;
            //URIを通常のパス形式に変換する
            Uri u = new Uri(path);
            path = u.LocalPath + Uri.UnescapeDataString(u.Fragment);
            return System.IO.Path.GetDirectoryName(path);
        }

    }

}
