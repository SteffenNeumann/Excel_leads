on FetchMessages(scriptText)
	try
		return run script scriptText
	on error errMsg number errNum
		return "ERROR:" & errNum & ":" & errMsg
	end try
end FetchMessages

on CopyFile(params)
	-- params: "sourcePath|destPath"
	set delimPos to offset of "|" in params
	if delimPos is 0 then return "ERROR:-1:missing delimiter"
	set srcPath to text 1 thru (delimPos - 1) of params
	set dstPath to text (delimPos + 1) thru -1 of params
	try
		do shell script "cp " & quoted form of srcPath & " " & quoted form of dstPath
		return "OK"
	on error errMsg number errNum
		return "ERROR:" & errNum & ":" & errMsg
	end try
end CopyFile

on RemoveXattr(folderPath)
	try
		do shell script "xattr -rd com.apple.quarantine " & quoted form of folderPath
		return "OK"
	on error errMsg number errNum
		return "ERROR:" & errNum & ":" & errMsg
	end try
end RemoveXattr
