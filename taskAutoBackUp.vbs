'���и��������{�� = 'AAA
' �����޸ġ��Լ������޸ġ����ļ���׃���������ļ����{�����ɣ����ÿ�����������؄e��̎ ���Բ������_�^����׃���г���
set ws  = createObject("wscript.shell")
set fso = createObject("scripting.fileSystemObject")
set targetPath = fso.getfolder("") '�������

dim Names(3) 'AAA
Names(0) = ""  '�������e.g.��Notes
Names(1) = ""  '�������
Names(2) = ""  '�������

dim Objects(3) 'AAA
Objects(0) ="" '�������
Objects(1) ="" '�������
Objects(2) ="" '�������


set allFiles = targetPath.Files
fileNum = 0
for each file in allFiles
  fileNum =fileNum+1
Next  
if day(date) = 1 then
  fso.deleteFolder("C:\backup\"&month(date)-2) '�����޸ġ�
  fso.createFolder("C:\backup\"&month(date))  �����޸ġ�
end if
i=0
Do until i = 3  'AAA
temp =  "C:\backup\"&month(date) & "\" &Names(i) '�����޸ġ�
fso.CopyFile Objects(i), temp
fso.movefile temp, "C:\backup\" &month(date) & "\" & Names(i) & day(date) & ".txt" '�����޸ġ�
i = i+1
Loop