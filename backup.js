//
//  ワンクリックで指定したファイルのバックアップを取るjs
//  名づけて「backup.js」
//
//--- 変数を二つ、設定すること---------------------------------
//
// バックアック先フォルダ（jsのあるフォルダの直下）を設定
//
var backupD = "backup";
//
// バックアックするファイルを設定
//
var targetFile = "backup.js";
//
//------------------------------------------------------------
//  Shell関連の操作を提供するオブジェクトを取得
var sh = new ActiveXObject( "WScript.Shell" );
//  ファイル関連の操作を提供するオブジェクトを取得
var fs = new ActiveXObject( "Scripting.FileSystemObject" );
// カレントフォルダ（jsのあるフォルダ）を取得
var CurrentD = sh.CurrentDirectory +"\\";
// バックアック先フォルダのフルパスを生成
var targetFP = CurrentD+backupD+"\\";

//------------------------------------------------------------
//　バックアップ先フォルダの準備
//  ⇒指定したフォルダが存在してるかを確認する。
// 　 https://jscript.zouri.jp/Source/FileFolderCtrl.html
if( fs.FolderExists(targetFP) ){
    //あるときは何もしない
}else{
    //なければフォルダーを作る
    fs.CreateFolder( targetFP )
    WScript.Echo( "バックアップフォルダを作成しました。");
};
//------------------------------------------------------------
// YYMMDDHHmm_型で日時を取得
// ⇒https://goat-inc.co.jp/blog/2584/
var dt  = new Date();
 var ymdhm =         ("00" + (dt.getFullYear())).slice(-2);
 var ymdhm = ymdhm + ("00" + (dt.getMonth()+1)).slice(-2);
 var ymdhm = ymdhm + ("00" + (dt.getDate())).slice(-2);
 var ymdhm = ymdhm + ("00" + (dt.getHours())).slice(-2);
 var ymdhm = ymdhm + ("00" + (dt.getMinutes())).slice(-2);
 var ymdhm = ymdhm + "_";

//
//　ファイルコピー
fs.CopyFile( targetFile, targetFP+ymdhm+targetFile);

//　
WScript.Echo( "バックアップしました。⇒ .\\" + backupD + "\\" + ymdhm + targetFile );

//  オブジェクトを解放
fs = null;
sh = null;
