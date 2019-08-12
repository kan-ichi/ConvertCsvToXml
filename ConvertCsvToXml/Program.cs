using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;

namespace ConvertCsvToXml
{
    class Program
    {
        #region エントリーポイント

        /// <summary>
        /// エントリーポイント
        /// </summary>
        public static void Main(string[] _args)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;

            Request request = new Request();

            request.InputCsvFolderName = ConfigurationManager.AppSettings["InputCsvFolderName"];
            request.InputCsvEncoding = ConfigurationManager.AppSettings["InputCsvEncoding"];
            request.InputCsvSkipFirstRow = (ConfigurationManager.AppSettings["InputCsvSkipFirstRow"] == "Yes");
            request.OutputXmlFolderName = ConfigurationManager.AppSettings["OutputXmlFolderName"];
            request.OutputXmlFileName = ConfigurationManager.AppSettings["OutputXmlFileName"];
            request.OutputXmlEncoding = ConfigurationManager.AppSettings["OutputXmlEncoding"];

            request.XsdFilePath = Path.Combine(path, "ConvertCsvToXml.exe.xsd");

            request.InputCsvFileNameByArgs = null;
            request.OutputXmlFileNameByArgs = null;

            if (_args.Length >= 1)
            {
                request.InputCsvFileNameByArgs = _args[0];
            }

            if (_args.Length >= 2)
            {
                request.OutputXmlFileNameByArgs = _args[1];
            }

            MainProcess(request);
        }

        #endregion

        #region リクエスト変数

        /// <summary>
        /// リクエスト変数
        /// </summary>
        public struct Request
        {
            /// <summary>
            /// 入力CSVフォルダ名
            /// </summary>
            public string InputCsvFolderName { get; set; }

            /// <summary>
            /// 入力CSVエンコーディング
            /// </summary>
            public string InputCsvEncoding { get; set; }

            /// <summary>
            /// 入力CSVの最初の一行を読込対象外とするか？
            /// </summary>
            public bool InputCsvSkipFirstRow { get; set; }

            /// <summary>
            /// 出力XMLフォルダ名
            /// </summary>
            public string OutputXmlFolderName { get; set; }

            /// <summary>
            /// 出力XMLファイル名
            /// </summary>
            public string OutputXmlFileName { get; set; }

            /// <summary>
            /// 出力XMLエンコーディング
            /// </summary>
            public string OutputXmlEncoding { get; set; }

            /// <summary>
            /// XMLスキーマファイルのパス
            /// </summary>
            public string XsdFilePath { get; set; }

            /// <summary>
            /// コマンドライン引数で指定された入力CSVファイル名
            /// </summary>
            public string InputCsvFileNameByArgs { get; set; }

            /// <summary>
            /// コマンドライン引数で指定された出力XMLファイル名
            /// </summary>
            public string OutputXmlFileNameByArgs { get; set; }
        }

        #endregion

        #region メイン処理

        /// <summary>
        /// メイン処理
        /// </summary>
        private static void MainProcess(Request _request)
        {
            WriteConsoleLogMessage("CSV -> XML 変換処理を開始します");

            // 処理する入力ファイルのリストを作成
            // コマンドライン引数で入力ファイル名が指定されていないならば、
            // アプリケーション構成ファイル定義のフォルダ内のすべてのファイルを処理
            // 入力ファイル名がコマンドライン引数で指定されているのならば、そのファイルだけを処理
            FileInfo[] inputFiles;
            if (string.IsNullOrEmpty(_request.InputCsvFileNameByArgs))
            {
                DirectoryInfo di = new DirectoryInfo(_request.InputCsvFolderName);
                inputFiles = di.GetFiles("*.csv", System.IO.SearchOption.TopDirectoryOnly);
                Array.Sort<FileInfo>(inputFiles, delegate (FileInfo a, FileInfo b)
                {
                    return a.Name.CompareTo(b.Name);
                });
            }
            else
            {
                inputFiles = new FileInfo[1] { new FileInfo(_request.InputCsvFileNameByArgs) };
            }

            // 出力ファイルについて
            // 引数で出力ファイル名が指定されていれば、それを使用
            // 指定されていないのならば、アプリケーション構成ファイル定義の値を使用
            string outputXmlFolderAndFileName;
            outputXmlFolderAndFileName = _request.OutputXmlFileNameByArgs ?? Path.Combine(_request.OutputXmlFolderName, _request.OutputXmlFileName);

            // 出力XMLのスキーマ情報に従い、出力用データテーブルを作成
            DataTable outputXmlTable = new DataTable();
            {
                outputXmlTable.ReadXmlSchema(_request.XsdFilePath);
            }

            // 入力ファイルを読み、出力用データテーブルに値を設定する
            foreach (var inputFile in inputFiles)
            {
                if (inputFile.Extension.ToLower() != ".csv") continue; // 「.csv2」などの拡張子が検索されるが、それは処理対象外とする
                WriteConsoleLogMessage("読み込んでいます " + inputFile.FullName);

                var parser = new TextFieldParser(inputFile.FullName, Encoding.GetEncoding(_request.InputCsvEncoding));
                using (parser)
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(","); // カンマ区切り

                    parser.HasFieldsEnclosedInQuotes = true; // フィールドが引用符で囲まれているか
                    parser.TrimWhiteSpace = false; // フィールドの空白トリム設定

                    long lineNumber = 0;
                    while (!parser.EndOfData) // ファイルの終端までループ
                    {
                        string[] row = parser.ReadFields(); // フィールドを読込
                        lineNumber++;
                        if (_request.InputCsvSkipFirstRow && lineNumber == 1) continue; // 最初の一行を読込対象外としている場合、読み飛ばし

                        DataRow dr = outputXmlTable.NewRow();
                        for (int i = 0; i < outputXmlTable.Columns.Count; i++)
                        {
                            if (row.Length > i)
                            {
                                dr[i] = row[i];
                            }
                            else // CSVの列が不足している場合、その列は空白とする
                            {
                                dr[i] = string.Empty;
                            }
                        }
                        outputXmlTable.Rows.Add(dr);
                    }
                }
            }

            WriteConsoleLogMessage("作成しています " + outputXmlFolderAndFileName);

            // 出力用データテーブルから、XMLファイルを作成する
            XmlTextWriter writer;
            {
                Encoding enc = Encoding.GetEncoding(_request.OutputXmlEncoding);
                writer = new XmlTextWriter(outputXmlFolderAndFileName, enc);
                writer.Formatting = Formatting.Indented;
                writer.WriteStartDocument();
                outputXmlTable.WriteXml(writer);
                writer.WriteEndDocument();
                writer.Close();
            }

            WriteConsoleLogMessage("CSV -> XML 変換処理を終了します");
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// コンソールにメッセージを出力します
        /// </summary>
        private static void WriteConsoleLogMessage(string _message)
        {
            Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " " + _message);
        }

        #endregion
    }
}
