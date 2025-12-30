#Requires AutoHotkey v2
#Include Lib\UIA.ahk

Main() {
	SplitPath(A_ScriptName, , , , &scriptName)
	PATH_INI_FILE := scriptName . ".ini"
	INI_SECTION_COMMON := "Common"
	INI_SECTION_LOG := "Log"
	INI_SECTION_TARGET_WINDOW := "TargetWindow"
	INI_SECTION_TARGET_ELEMENT := "TargetElement"

	checkInterval := Integer(IniRead(PATH_INI_FILE, INI_SECTION_COMMON, "CheckInterval"))
	invalidResultText := IniRead(PATH_INI_FILE, INI_SECTION_COMMON, "InvalidResultText")

	logDirectoryPath := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "DirectoryPath")
	logFilenamePrefix := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "FilenamePrefix")
	logEncoding := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "Encoding")
	logVersion := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "Version")
	logMetaInfo := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "MetaInfo", "")
	logTimestampFormat := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "TimestampFormat", "")
	logSessionIdFormat := IniRead(PATH_INI_FILE, INI_SECTION_LOG, "SessionIdFormat", "")

	targetWindowTitle := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_WINDOW, "WindowTitle")
	targetWindowText := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_WINDOW, "WindowText")
	targetWindowCheckTimeout := Integer(IniRead(PATH_INI_FILE, INI_SECTION_TARGET_WINDOW, "CheckTimeout"))

	sourceElementPath := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_ELEMENT, "SourcePath")
	resultElementPath := IniRead(PATH_INI_FILE, INI_SECTION_TARGET_ELEMENT, "ResultPath")

	targetWindowHwnd := WinWait(targetWindowTitle, targetWindowText, targetWindowCheckTimeout)
	if (!targetWindowHwnd) {
		MsgBox("Window is not found.")
		ExitApp
	}
	targetWindow := UIA.ElementFromHandle(targetWindowHwnd)

	sourceElement := targetWindow.ElementFromPathExist(sourceElementPath)
	resultElement := targetWindow.ElementFromPathExist(resultElementPath)
	if (!sourceElement || !resultElement) {
		MsgBox("Element is not found.")
		ExitApp
	}

	lastResultText := resultElement.Value
	Loop {
		if (!WinExist(targetWindowTitle, targetWindowText)) {
			ExitApp
		}
		
		sourceElement := targetWindow.ElementFromPathExist(sourceElementPath)
		resultElement := targetWindow.ElementFromPathExist(resultElementPath)
		if (sourceElement && resultElement) {
			currentResultText := resultElement.Value
			if (IsValidText(invalidResultText, currentResultText) && currentResultText != lastResultText) {
				sourceText := sourceElement.Value
				lastResultText := currentResultText
				
				timestamp := FormatTime(, logTimestampFormat)
				sessionId := FormatTime(, logSessionIdFormat)
				
				logMetaInfoStr := ((StrLen(logMetaInfo) > 0) ? ("`"info`": " . FormatJsonString(logMetaInfo) . ",") : "")
				
				json :=
					"{"
						. "`"timestamp`": " . FormatJsonString(timestamp) . ","
						. "`"text`": {"
							. "`"source`": " . FormatJsonString(sourceText) . ","
							. "`"result`": " . FormatJsonString(currentResultText) . ","
							. "`"source_length`": " . StrLen(sourceText) . ","
							. "`"result_length`": " . StrLen(currentResultText)
						. "},"
						. "`"meta`": {"
							. "`"version`": " . FormatJsonString(logVersion) . ","
							. "`"session_id`": " . FormatJsonString(sessionId)
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

EscapeString(str) {
	str := StrReplace(str, "\", "\\")
	str := StrReplace(str, "`"", "\`"")
	str := StrReplace(str, "`n", "\n")
	str := StrReplace(str, "`r", "\r")
	return str
}

IsValidText(invalidText, text) {
	return (text != "" && EscapeString(text) != invalidText)
}

FormatJsonString(str) {
	return "`"" . EscapeString(str) . "`""
}

Main()
