'若有更改需作調整 = 'AAA
' 【待修改】以及【待修改】即文件名變量，根據文件名調整即可（因為每個基本都有特別用處 所以不能在開頭就以變量列出）
set ws  = createObject("wscript.shell")
set fso = createObject("scripting.fileSystemObject")
set targetPath = fso.getfolder("") '【待填寫】

dim Names(3) 'AAA
Names(0) = ""  '【待填寫】e.g.：Notes
Names(1) = ""  '【待填寫】
Names(2) = ""  '【待填寫】

dim Objects(3) 'AAA
Objects(0) ="" '【待填寫】
Objects(1) ="" '【待填寫】
Objects(2) ="" '【待填寫】


set allFiles = targetPath.Files
fileNum = 0
for each file in allFiles
  fileNum =fileNum+1
Next  
if day(date) = 1 then
  fso.deleteFolder("C:\backup\"&month(date)-2) '【待修改】
  fso.createFolder("C:\backup\"&month(date))  【待修改】
end if
i=0
Do until i = 3  'AAA
temp =  "C:\backup\"&month(date) & "\" &Names(i) '【待修改】
fso.CopyFile Objects(i), temp
fso.movefile temp, "C:\backup\" &month(date) & "\" & Names(i) & day(date) & ".txt" '【待修改】
i = i+1
Loop