' Sub process to append BOM to File
Sub appendBOM(msg)
Dim input
Set input = CreateObject("ADODB.Stream")
input.Type = 2    ' 1：Binary・2：Text
input.Charset = "UTF-8"    ' apply Charset
input.Open    ' Open Stream object
input.LoadFromFile msg   ' Load file

'Override File
Dim output
Set output = CreateObject("ADODB.Stream")
output.Type = 2
output.Charset = "UTF-8"
output.Open

' Load file and write 
Dim records
Do Until input.EOS
  Dim lineStr
  lineStr = input.ReadText(-2)    ' -1：Load all line・-2：Load one line
  output.WriteText lineStr, 1    ' 0：Write String・1：Write String + new line
Loop

' Save out put
output.SaveToFile msg, 2    '1：Save as new file・2：Override File

' Close Stream 
input.Close
output.Close
End Sub

Dim msg
    Set objArgs = WScript.Arguments 'Prepare Drag and Drop object
      For I = 0 to objArgs.Count - 1 'Repeat for all files dropped in
      msg = objArgs(I) 
      Call appendBOM(msg) 
Next
