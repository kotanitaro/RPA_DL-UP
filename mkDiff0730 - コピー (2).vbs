'ファイル名を指定する変数を定義
Dim strDate
Dim strDateYesterday
Dim strDateYesterdaySlashed
Dim cmdline


'今日の日付を取得
strDateSlashed = Date
strDate = Replace(strDateSlashed,  "/","")
'昨日の日付を取得
strDateYesterdaySlashed = DateAdd("d",-1,strDateSlashed)
strDateYesterday = Replace(strDateYesterdaySlashed,  "/","")

cmdline = "fc.exe C:\medicFileList\"+strDateYesterday+"2300.txt C:\medicFileList\"+strDate+"0730.txt"

'ファイル処理の変数を定義
Dim objShell
Dim objExec
Dim objFSO
Dim objFile
Dim strResult

'ファイル処理を実行

'　差分チェック
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec(cmdline)

Do While objExec.Status = 0
   WScript.Sleep 100
Loop

'　差分チェック結果を保存
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\medicFileList\diff"+strDate+"0730.csv",8,True)
strResult = cStr(objExec.StdOut.ReadAll)
objFile.WriteLine(strResult)
objFile.Close
Set objFile = Nothing
Set objFSO = Nothing