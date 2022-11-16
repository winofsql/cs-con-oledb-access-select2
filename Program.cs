using System;
using System.Data.OleDb;

namespace cs_con_oledb_access_select2
{
    class Program
    {
        static void Main(string[] args)
        {
            OleDbConnection myCon;
            OleDbCommand myCommand;
            OleDbDataReader myReader;

            using (myCon = new OleDbConnection())
            using (myCommand = new OleDbCommand())
            {
                // SQL文字列格納用
                string myQuery = "";
                string myPath = @"C:\app\workspace\販売管理.accdb";
                // string myPath = @"C:\app\workspace\販売管理.mdb";	// 古い MsAccess

                // 接続文字列の作成
                myCon.ConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};", myPath);
                // 出力ウインドウに表示
                // デバッグ>オプション の 『出力』で、【全てのデバッグ出力】以外を オフにする
                Console.WriteLine(myCon.ConnectionString);

                // *********************
                // 接続
                // *********************
                try {
                    // 接続文字列を使用して接続
                    myCon.Open();
                    // コマンドオブジェクトに接続をセット
                    myCommand.Connection = myCon;
                    // コマンドを通常 SQL用に変更
                    myCommand.CommandType = System.Data.CommandType.Text;
                }
                catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                    return;
                }

                myQuery = "select * from 社員マスタ";
                myCommand.CommandText = myQuery;

                // *********************
                // レコードセット取得
                // *********************
                try {
                    using (myReader = myCommand.ExecuteReader())
                    {
                        // *********************
                        // 列数
                        // *********************
                        int nCols = myReader.FieldCount;
                        Type fldType;

                        // カラムループ用
                        int idx = 0;
                        while (myReader.Read()) {

                            for (idx = 0; idx <= nCols - 1; idx++) {
                                if (idx != 0) {
                                    Console.Write(",");
                                }

                                // NULL でない場合
                                if (!myReader.IsDBNull(idx)) {
                                    // 列のデータ型を取得
                                    fldType = myReader.GetFieldType(idx);

                                    // 文字列
                                    if (fldType.Name == "String") {
                                        Console.Write(myReader.GetValue(idx) + "");
                                        continue;
                                    }
                                    if (fldType.Name == "Int32") {
                                        Console.Write(myReader.GetInt32(idx).ToString() + "");
                                        continue;
                                    }
                                    if (fldType.Name == "DateTime") {
                                        Console.Write(myReader.GetDateTime(idx).ToString() + "");
                                        continue;
                                    }

                                    Console.Write(myReader.GetValue(idx).ToString() + "");

                                }
                                else {
                                    Console.Write("");
                                }
                            }
                            // 1行の最後
                            Console.WriteLine("");

                        }

                    }
                }
                catch (Exception ex) {
                    myCon.Close();
                    Console.WriteLine(ex.Message);
                    return;
                }

            }
        }
    }
}
