'~~~~~~~~~~~~~~~~~~~|Create file in user Desktop| [test 5001830]~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'For randomize
Dim min, max
min = 1
max = 100
userPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%\Desktop")
Randomize 'everytime random value
rndNumber = Int((max-min+1)*Rnd+min)
file = "\BIGTEST_VBS_" & rndNumber
set fs = CreateObject("Scripting.FileSystemObject")
set drop = fs.CreateTextFile(userPath & file, false)
drop.Write("testString")
drop.Close

'MsgBox("File : [ " & file & " ] created on Desktop, go CALC!")

'~~~~~~~~~~~~~~~~|Execute some shit (calc)| [test 5000043]~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CreateObject("WScript.Shell").Run("calc.exe")

'~~~~~~~~~~~~~~~~|External network connection| [test 5001114]~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set connect = CreateObject("Microsoft.XMLHTTP")
connect.open "GET", "http://checkip.dyndns.org", False
connect.send
WScript.Echo "OPAOPA_EXTERNAL_IP:" & connect.responseText
