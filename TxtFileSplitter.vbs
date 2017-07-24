'把原始文件的名称改为“0.txt”，然后把本程序放到原始文件所在文件夹
'然后双击。n 为要要分割的文件的行数。
set g_fso = CreateObject("Scripting.FileSystemObject")
dim n 
n = 1000
dim count 
count = 0
dim fileCount
fileCount = 1
set objStream = g_fso.openTextFile(".\0.txt")
Set tarStream = g_fso.OpenTextFile(".\1.txt", 2, True)

Do While objStream.AtEndOfStream <> True

	strLine = objStream.ReadLine
	tarStream.writeline strLine
	count = count + 1
	
	if count MOD n  = 0 then
		fileCount = fileCount + 1
		tarStream.close
		Set tarStream = g_fso.OpenTextFile(".\" & fileCount & ".txt", 2, True)
	end if
	
Loop

objStream.close

msgbox "处理完毕！"
