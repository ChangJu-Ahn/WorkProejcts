<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b21mb2.asp
'*  4. Program Name         : 사양항목 등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
Call LoadBasisGlobalInf


Dim oPB3S211																	'☆ : 저장용 ComProxy Dll 사용 변수 

Dim StrNextKey											'⊙: 다음 값 
Dim lgStrPrevKey										'⊙: 이전 값 
Dim strMode												'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngMaxRow											'⊙: 현재 그리드의 최대Row
Dim intFlgMode
Dim LngRow
Dim LngIdx

Dim I1_B_Char
Dim IG1_B_Char_Value

Dim arrCols, arrRows									'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim strStatus											'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
Dim	lGrpCnt												'☜: Group Count
Dim strCode												'⊙: Lookup 용 리턴 변수 

ReDim I1_B_Char(3)
Const C_CommandSent = 0
Const C_Char_Cd = 1
Const C_Char_Nm = 2
Const C_Char_Value_Digit = 3

'Export Group
Const C_IG1_Select_Char = 0
Const C_IG1_Char_Value_Cd = 1
Const C_IG1_Char_Value_Nm = 2

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")											'☜ : 현재 상태를 받음 
	intFlgMode = CInt(Request("txtFlgMode"))	
	LngMaxRow = CInt(Request("txtMaxRows"))									'☜: 최대 업데이트된 갯수 
    lgStrPrevKey = Trim(Request("lgStrPrevKey"))

    If intFlgMode = OPMD_CMODE Then
		I1_B_Char(C_CommandSent)	= "CREATE"
    Else
		I1_B_Char(C_CommandSent)	= "UPDATE"
    End If

    I1_B_Char(C_Char_Cd)			= UCase(Request("txtCharCd1"))
	I1_B_Char(C_Char_Nm)			= Request("txtCharNm1")
	I1_B_Char(C_Char_Value_Digit)	= Request("txtCharValueDigit")

    If I1_B_Char(C_Char_Cd) = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If
	
	If I1_B_Char(C_Char_Nm) = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If
	
	If I1_B_Char(C_Char_Value_Digit) = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	ElseIf I1_B_Char(C_Char_Value_Digit) < 1 Or I1_B_Char(C_Char_Value_Digit) > 16 Then
		Call DisplayMsgBox("122645", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If

	arrRows = Split(Request("txtSpread"), gRowSep)							'☆: Spread Sheet 내용을 담고 있는 Element명 
    ReDim IG1_B_Char_Value(UBound(arrRows,1),2)

	For LngRow = 0 To LngMaxRow - 1
		arrCols = Split(arrRows(LngRow), gColSep)
		strStatus = UCase(Trim(arrCols(0)))									'☜: Row 의 상태 

		Select Case strStatus
			
		    Case "C", "U"
				If IsVaildCodeLength(Trim(arrCols(2)), I1_B_Char(C_Char_Value_Digit)) = False Then
					Call DisplayMsgBox("122649", vbOKOnly, "", "", I_MKSCRIPT)
					Response.End
				End If

				IG1_B_Char_Value(LngRow, C_IG1_Select_Char)		= strStatus
				IG1_B_Char_Value(LngRow, C_IG1_Char_Value_Cd)	= UCase(Trim(arrCols(2)))
				IG1_B_Char_Value(LngRow, C_IG1_Char_Value_Nm)	= arrCols(3)
			Case "D"
				IG1_B_Char_Value(LngRow, C_IG1_Select_Char)		= "D"
				IG1_B_Char_Value(LngRow, C_IG1_Char_Value_Cd)	= UCase(Trim(arrCols(2)))
	
		End Select
	Next
	
	Set oPB3S211 = Server.CreateObject("PB3S211.cBMngChar")

	If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

	Call oPB3S211.B_MANAGE_CHAR(gStrGlobalCollection, _
								I1_B_Char, _
								IG1_B_Char_Value)

	If CheckSYSTEMError(Err,True) = True Then
       Set oPB3S211 = Nothing		                                        '☜: Unload Comproxy DLL
       Response.End		
    End If

	Set oPB3S211 = Nothing                                                  '☜: Unload Comproxy

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr																
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr					
	
	'==============================================================================
	' Function : SheetFocus
	' Description : 에러발생시 Spread Sheet에 포커스줌 
	'==============================================================================
	Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
		Dim strHTML
		If iLoc = I_INSCRIPT Then
			strHTML = "parent.frm1.vspdData.focus" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
			Response.Write strHTML
		ElseIf iLoc = I_MKSCRIPT Then
			strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
			strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
			Response.Write strHTML
		End If
	End Function

	'==============================================================================
	' Function : IsVaildCodeLength (User Defined Function)
	' Description : 문자열의 길이를 원하는 길이와 비교하여 작거나 같으면 True 리턴 
	'==============================================================================
	Function IsVaildCodeLength(Byval iStr, Byval iDigit)
		Dim intLength
		Dim intIdx
		Dim intAsc
		Dim intSum
		
		IsVaildCodeLength = True
		
		intSum = 0
		intLength = Len(iStr)
		
		For intIdx=0 To intLength-1
			intAsc = ASC(Mid(iStr,intIdx+1,1))
			If CInt(intAsc) < 0 Or CInt(intAsc) > 255 Then
				intSum = intSum + 2
			Else
				intSum = intSum + 1
			End If
		Next
		
		If intSum > CInt(iDigit) Then
			IsVaildCodeLength = False
		End If
	End Function
%>
