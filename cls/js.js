var cls = function(){

    this.fnDo = function(){
        WScript.Echo("test中");
        return new cls
    }; 

};

var a = new cls()

a.fnDo();git