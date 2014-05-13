<SCRIPT RUNAT=Server LANGUAGE="VBScript">
Function DanDate (strDate, strFormat)

	Dim intPosItem
	Dim intHourPart
	Dim strHourPart
	Dim strMinutePart
	Dim strSecondPart
	Dim strAMPM

	
	If not IsDate(strDate) Then
		DanDate = strDate
		Exit Function
	End If
	
	intPosItem = Instr(strFormat, "%m")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("m",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%m")
	Loop

	intPosItem = Instr(strFormat, "%b")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        MonthName(DatePart("m",strDate),True) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%b")
	Loop
	
	intPosItem = Instr(strFormat, "%B")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        MonthName(DatePart("m",strDate),False) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%B")
	Loop
	
	intPosItem = Instr(strFormat, "%d")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("d",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%d")
	Loop

	intPosItem = Instr(strFormat, "%j")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("y",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%j")
	Loop

	intPosItem = Instr(strFormat, "%y")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        Right(DatePart("yyyy",strDate),2) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%y")
	Loop

	intPosItem = Instr(strFormat, "%Y")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("yyyy",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%Y")
	Loop

	intPosItem = Instr(strFormat, "%w")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("w",strDate,1) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%w")
	Loop

	intPosItem = Instr(strFormat, "%a")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        WeekDayName(DatePart("w",strDate,1),True) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%a")
	Loop

	intPosItem = Instr(strFormat, "%A")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        WeekDayName(DatePart("w",strDate,1),False) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%A")
	Loop

	intPosItem = Instr(strFormat, "%I")
	Do While intPosItem > 0
		intHourPart = DatePart("h",strDate) mod 12
		if intHourPart = 0 then intHourPart = 12
                strFormat = Left(strFormat, intPosItem-1) & _
                        intHourPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%I")
	Loop

	intPosItem = Instr(strFormat, "%H")
	Do While intPosItem > 0
		strHourPart = DatePart("h",strDate)
		if strHourPart < 10 Then strHourPart = "0" & strHourPart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strHourPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%H")
	Loop

	intPosItem = Instr(strFormat, "%M")
	Do While intPosItem > 0
		strMinutePart = DatePart("n",strDate)
		if strMinutePart < 10 then strMinutePart = "0" & strMinutePart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strMinutePart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%M")
	Loop

	intPosItem = Instr(strFormat, "%S")
	Do While intPosItem > 0
		strSecondPart = DatePart("s",strDate)
		if strSecondPart < 10 then strSecondPart = "0" & strSecondPart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strSecondPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%S")
	Loop

	intPosItem = Instr(strFormat, "%P")
	Do While intPosItem > 0
		if DatePart("h",strDate) >= 12 then
			strAMPM = "PM"
		Else
			strAMPM = "AM"
		End If
                strFormat = Left(strFormat, intPosItem-1) & _
                        strAMPM & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%P")
	Loop

	intPosItem = Instr(strFormat, "%%")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & "%" & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%%")
	Loop

	DanDate = strFormat
	
End Function
</SCRIPT>

