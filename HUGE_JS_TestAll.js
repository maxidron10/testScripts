//~~~~~|Create file in user Desktop| [test 5001830]~~~~~~~~~~
var file = "\\BIGTEST_JS_" + Math.floor(Math.random() * 101);
var shell = new ActiveXObject("WScript.Shell");
var fs = new ActiveXObject("Scripting.FileSystemObject");
var userDesktop = shell.SpecialFolders("Desktop");
var drop = fs.CreateTextFile(userDesktop + file, true);
drop.Write("testString");
drop.Close();

//shell.Popup("File: [ " + file + " ] on Desktop, go MSPAINT!");

//~~~~~~|Execute some shit (calc)| [test 5000043]~~~~~~~~~~~~
shell.Run("calc.exe");

//~~~~~|External network connection| [test 5001114]~~~~~~~~~~
var req = new ActiveXObject("Microsoft.XMLHTTP"); 
req.open('GET', "http://checkip.dyndns.org", false);   
req.send(null);  
if(req.status == 200)
{
	var show = new ActiveXObject("WScript.Shell");
	show.Popup("OPAOPA_EXTERNAL_IP:"+ req.responseText);
}
