#Requires AutoHotkey v2
#Include Lib\UIA.ahk

Main() {
	MSG_LOGGING_STARTED := "Logging started."
	MSG_LOGGING_STOPPED := "Logging stopped."
	
	ERROR_MSG_WINDOW_IS_NOT_FOUND := "Window is not found."
	ERROR_MSG_ELEMENT_IS_NOT_FOUND := "Element is not found."
	
	SEPARATOR_PAIR := ":"
	
	SplitPath(A_ScriptName, , , , &scriptName)
	PATH_INI_FILE := scriptName . ".ini"
	INI_SECTION_COMMON := "Common"
	INI_SECTION_LOG := "Log"
	INI_SECTION_TARGET_WINDOW := "TargetWindow"
	INI_SECTION_TARGET_ELEMENT := "TargetElement"

	checkInterval := Integer(IniRead(PATH_INI_FILE, INI_SECTION_COMMON, "CheckInterval"))
	invalidEscapedResultTexts := ParseArrayString(IniRead(PATH_INI_FILE, INI_SECTION_COMMON, "InvalidEscapedResultTexts"))

	logDirectoryPath := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "DirectoryPath")
	logFilenamePrefix := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "FilenamePrefix")
	logEncoding := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "Encoding")
	logVersion := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "Version")
	logMetaInfo := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "MetaInfo", "")
	logTimestampFormat := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "TimestampFormat", "")
	logSessionIdFormat := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "SessionIdFormat", "")
	
	targetWindowTitle := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_WINDOW, "WindowTitle")
	targetWindowText := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_WINDOW, "WindowText")
	
	sourceElementPath := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_ELEMENT, "SourcePath")
	resultElementPath := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_ELEMENT, "ResultPath")
	etcElementPaths := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_ELEMENT, "EtcPaths", "[]")
	
	etcElementPathMap := Map()
	for (pairStr in ParseArrayString(etcElementPaths)) {
		pair := StrSplit(pairStr, SEPARATOR_PAIR)
		if (pair.Length >= 2) {
			tempKey := pair[1]
			tempValue := pair[2]
			if (StrLen(tempKey) > 0 && StrLen(tempValue) > 0) {
				etcElementPathMap[tempKey] := tempValue
			}
		}
	}
	
	targetWindowHwnd := WinExist(targetWindowTitle, targetWindowText)
	if (!targetWindowHwnd) {
		MsgBox(ERROR_MSG_WINDOW_IS_NOT_FOUND . "`n" . MSG_LOGGING_STOPPED)
		ExitApp
	}
	targetWindow := UIA.ElementFromHandle(targetWindowHwnd)
	
	sourceElement := targetWindow.ElementFromPathExist(sourceElementPath)
	resultElement := targetWindow.ElementFromPathExist(resultElementPath)
	if (!sourceElement || !resultElement) {
		MsgBox(ERROR_MSG_ELEMENT_IS_NOT_FOUND . "`n" . MSG_LOGGING_STOPPED)
		ExitApp
	}
	
	lastResultText := resultElement.Value
	MsgBox(MSG_LOGGING_STARTED)
	Loop {
		if (!WinExist(targetWindowTitle, targetWindowText)) {
			MsgBox(ERROR_MSG_WINDOW_IS_NOT_FOUND . "`n" . MSG_LOGGING_STOPPED)
			ExitApp
		}
		
		sourceElement := targetWindow.ElementFromPathExist(sourceElementPath)
		resultElement := targetWindow.ElementFromPathExist(resultElementPath)
		if (sourceElement && resultElement) {
			currentResultText := GetText(resultElement)
			if (IsValidText(invalidEscapedResultTexts, currentResultText) && currentResultText != lastResultText) {
				lastResultText := currentResultText
				sourceText := GetText(sourceElement)
				
				timestamp := FormatTime(, logTimestampFormat)
				sessionId := FormatTime(, logSessionIdFormat)
				
				etcStr := ""
				For (tempKey, tempValue in etcElementPathMap) {
					tempElement := targetWindow.ElementFromPathExist(tempValue)
					if (tempElement) {
						tempText := GetText(tempElement)
						if (StrLen(etcStr) > 0) {
							etcStr .= ", "
						}
						etcStr .= "`"" . tempKey . "`": " . FormatJsonString(tempText)
					}
				}
				if (StrLen(etcStr) > 0) {
					etcStr := ", `"etc`": {" . etcStr . "}"
				}
				
				logMetaInfoStr := ((StrLen(logMetaInfo) > 0) ? (", `"info`": " . FormatJsonString(logMetaInfo)) : "")
				
				json :=
					"{"
						.   "`"timestamp`": " . FormatJsonString(timestamp)
						. ", `"text`": {"
							.   "`"source`": " . FormatJsonString(sourceText)
							. ", `"result`": " . FormatJsonString(currentResultText)
							. ", `"source_length`": " . StrLen(sourceText)
							. ", `"result_length`": " . StrLen(currentResultText)
						. "}"
						. etcStr
						. ", `"meta`": {"
							.   "`"version`": " . FormatJsonString(logVersion)
							. ", `"session_id`": " . FormatJsonString(sessionId)
							. logMetaInfoStr
						. "}"
					. "}"
				
				saveDirectoryPath := logDirectoryPath . "\" . FormatTime(, "yyyy") . "\" . FormatTime(, "MM")
				DirCreate(saveDirectoryPath)
				
				logFileName := logFilenamePrefix . FormatTime(, "yyyy-MM-dd") . ".json"
				FileAppend(json . "`n", saveDirectoryPath . "\" . logFileName, logEncoding)
			}
		}
		
		Sleep(checkInterval)
	}
}

GetText(element) {
	try {
		hasValuePattern := element.GetPattern(UIA.Pattern.Value)
	} catch {
		hasValuePattern := False
	}
	if (hasValuePattern) {
		str := element.Value
	} else {
		str := element.Name
	}
	return str
}

ParseArrayString(str) {
	resultArray := []
	
	if (!(RegExMatch(str, "^\s*\[(.*)\]\s*$", &m))) {
		return []
	}
	
	innerStr := m[1]
	
	if (StrLen(innerStr) > 0) {
		pos := 1
		while (RegExMatch(innerStr, '(?<!\\)"(.*?)(?<!\\)"', &m, pos)) {
			resultArray.Push(m[1])
			pos := m.Pos + m.Len
		}
	}
	
	return resultArray
}

EscapeString(str) {
	str := StrReplace(str, "\", "\\")
	str := StrReplace(str, "`"", "\`"")
	str := StrReplace(str, "`n", "\n")
	str := StrReplace(str, "`r", "\r")
	return str
}

IsValidText(invalidEscapedTexts, text) {
	result := True
	if (text != "") {
		escapedText := EscapeString(text)
		for (invalidEscapedText in invalidEscapedTexts) {
			if (escapedText == invalidEscapedText) {
				result :=  False
				break
			}
		}
	} else {
		result := False
	}
	return result
}

FormatJsonString(str) {
	return "`"" . EscapeString(str) . "`""
}

Main()
