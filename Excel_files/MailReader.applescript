on FetchMessages(scriptText)
	try
		return run script scriptText
	on error errMsg number errNum
		return "ERROR:" & errNum & ":" & errMsg
	end try
end FetchMessages
