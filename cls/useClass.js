
  var clsShell = function(){
    this.obj = new ActiveXObject("WScript.SHell");
  }
  var clsPCsetting = function(){
    this.net = new ActiveXObject("WScript.Network");
    this.she = new ActiveXObject("WScript.SHell");
  }
  
function get_prototype(cls){
  function Dummy(){}
  Dummy.prototype = cls.prototype
  return new Dummy();
}

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
  function Apploader (){
    clsShell.apply(this,arguments);
    this.chk = this.Chk;
    this.get = this.Get;
    this.run = this.Run;
    this.open = this.Open;
    this.check = this.Check;
  }
    Apploader.prototype = get_prototype(clsShell);
    Apploader.prototype.Get = function ( url ){
      return this.obj.Run("\"%comspec% /c " + url + "\"", 1, true)
    };  
    Apploader.prototype.Run = function ( url ){
      return this.obj.Run("\"%comspec% /c " + url + "\"", 1, false)
    };
    Apploader.prototype.Script = function ( url ){
      return this.obj.Run("C:/WINDOWS/system32/wscript.exe \"" + url + "\"", 1, false)
    };
    Apploader.prototype.Open = function ( url ){
      return this.obj.Run( "\"" + url + "\"", 1, false)
    };
    Apploader.prototype.Check = function ( url ){
      return this.obj.Run( "\"" + url + "\"", 1, true)
    };

  function RegistrySetting ( root ){
    clsShell.apply(this,arguments);
    this.settext = this.setText = this.SetText;
    this.setint = this.setInt = this.SetInt;
    this.setsptext = this.setSpText = this.SetSpText;
    this.setbinary = this.setBinary = this.SetBinary;
  }
    RegistrySetting.prototype = get_prototype(clsShell);
    RegistrySetting.prototype.SetText = function ( key, word ){
      this.obj.RegWrite(root & key, word, "REG_SZ");
    };
    RegistrySetting.prototype.SetInt = function ( key, word ){
      this.obj.RegWrite(root & key, word, "REG_DWORD");
    };
    RegistrySetting.prototype.SetSpText = function ( key, word ){
      this.obj.RegWrite(root & key, word, "REG_EXPAND_SZ");
    };
    RegistrySetting.prototype.SetBinary = function ( key, word ){
      this.obj.RegWrite(root & key, word, "REG_BINARY");
    };
    RegistrySetting.prototype.GetValue = function ( key ){
      try {
        return this.obj.RegRead(root & key)
      } catch(e){
        if(e.number == -2147024894){
          return ""
        } else {
          return "error"
        }
      }
    }

  function PC (){
    clsPCsetting.apply(this);
    this.computername = this.ComputerName;
    this.UserDomain = this.userDomain;
    this.UserName = this.userName;
    this.Cmd = this.cmd;
  }
    PC.prototype = get_prototype(clsPCsetting);
    PC.prototype.ComputerName = function (){
      return this.net.ComputerName;
    };
    PC.prototype.userDomain = function (){
      return this.net.userDomain;
    };
    PC.prototype.userName = function (){
      return this.net.userName;
    };
    PC.prototype.cmd = function (code){
      return this.she.Exec('cmd /c '+code);
    };
    PC.prototype.ip = function (){
      var objIpSet
      var objIpConfig;
      var strComputer
      var objWMIService;
      var strip;

      strip = "";

      strComputer = ".";
      
      objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2");

      objIpSet = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE");

      for (objIpConfig in objIpSet) {
          if (!objIpConfig.IPAddress) {
              strip = objIpConfig.IPAddress[objIpConfig.IPAddress.lbound(1)];
              break;
          }
      }

      return strip;
    }

  var allWindow = function (){

    this.Pop = function ( text ){
      try {
        WScript.ECHO(text);
      } catch(e) {
        try {
          alert(text);
        } catch(e) {
          window.alert(text);
        }
      } 
    };

    this.alert = this.Alert = this.say = this.Say = this.pop = this.Pop;

  }
  
var my = function() {
  this.TextReader = function(fileName){return new TextReader(fileName)};
  this.Apploader = function(){return new Apploader()};
  this.RegistrySetting = function(root){return new RegistrySetting(root)};
  this.allWindow = function(){return new allWindow()};
  this.PC = function(){return new PC()};
}

