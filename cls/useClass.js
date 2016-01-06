var TextReader = function( fileName ){

  //  ファイル関連の操作を提供するオブジェクトを取得
  var fs;
  fs = new ActiveXObject( "Scripting.FileSystemObject" );

  //  オープンモード
  var FORREADING;    // 読み取り専用
  var FORWRITING;    // 書き込み専用
  var FORAPPENDING;  // 追加書き込み
  
  FORREADING      = 1;    // 読み取り専用
  FORWRITING      = 2;    // 書き込み専用
  FORAPPENDING    = 8;    // 追加書き込み

  //  開くファイルの形式
  var TRISTATE_TRUE;         // Unicode
  var TRISTATE_FALSE;        // ASCII
  var TRISTATE_USEDEFAULT;   // システムデフォルト
  
  TRISTATE_TRUE       = -1;   // Unicode
  TRISTATE_FALSE      =  0;   // ASCII
  TRISTATE_USEDEFAULT = -2;   // システムデフォルト

  var file;
  file = null;

  this.Read = function(){

                        var buf;
                        buf = "";

                        //  ファイルを読み取り専用で開く
                        clsOpen( fileName, FORREADING );

                        try {
                            if(!file.AtEndOfStream){
                            	buf = file.ReadAll();
			                      } else {
				                      buf = "";
			                      }
                        }
                        catch (e) {
                            buf = "error";
                        }

                        clsClose();

                        //  ファイルから全ての文字データを読み込む
                        return buf;
            };

  this.Write = function(text){
                        //  ファイルを読み取り専用で開く
                        clsOpen( fileName, FORWRITING );

                        //  ファイルへ全ての文字データを書き込む
                        file.Write( text );

                        return clsClose();
            };

  this.Clear = function(){
                        //  ファイルを読み取り専用で開く
                        clsOpen( fileName, FORWRITING );

                        //  ファイルへ全ての文字データを書き込む
                        file.Write( "" );

                        return clsClose();
            };

  this.WriteLine = function(text){
                        //  ファイルを読み取り専用で開く
                        clsOpen( fileName, FORAPPENDING );

                        //  ファイルへ一行文字データを読み込む
                        file.WriteLine( text );

                        return clsClose();
            };

  this.AddBlank = function(){
                        //  ファイルを読み取り専用で開く
                        clsOpen( fileName, FORAPPENDING );

                        //  ファイルへ一行改行を読み込む
                        file.WriteBlankLines( 1 );

                        return clsClose();

            };

  this.SearchWord = function( word ){
                        //  ファイルを読み取り専用で開く
                        clsOpen( fileName, FORREADING );

                        var buf;
                        buf = (file.ReadAll().match(new RegExp(word, "g")) || []).length;

                        clsClose();

                        //  ファイルから文字データを検索し該当する個数を取得する
                        return buf

            };

  function clsOpen( name , mode ){

                        if (!mode) { mode = FORREADING };

                        file = fs.OpenTextFile( name, mode, true, TRISTATE_FALSE );

            };

  function clsClose(){

                        file.Close();

                        return new TextReader( fileName );

            };

  this.RE = this.Re = this.re = this.read = this.Read;
  this.WR = this.Wr = this.wr = this.write = this.Write;
  this.WL = this.Wl = this.wl = this.writeline = this.writeLine = this.WriteLine;
  this.AB = this.Ab = this.ab = this.addblank = this.addBlank = this.AddBlank;
  this.CL = this.Cl = this.cl = this.clear = this.Clear;
  this.SW = this.Sw = this.sw = this.searchword = this.searchWord = this.SearchWord;

};

var Apploader = function (){

  var objShell = new ActiveXObject("WScript.Shell")

  this.Get = function ( url ){
			return objShell.Run("\"%comspec% /c " + url + "\"", 1, true)
            };

  this.Run = function ( url ){
			return objShell.Run("\"%comspec% /c " + url + "\"", 1, false)
            };

  this.Script = function ( url ){
			return objShell.Run("C:/WINDOWS/system32/wscript.exe \"" + url + "\"", 1, false)
            };

  this.Open = function ( url ){
			return objShell.Run( "\"" + url + "\"", 1, false)
            };

  this.Check = function ( url ){
			return objShell.Run( "\"" + url + "\"", 1, true)
            };

  this.chk = this.Chk
  this.get = this.Get;
  this.run = this.Run;
  this.open = this.Open;
  this.check = this.Check;

}