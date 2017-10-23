<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b21mb2.asp
'*  4. Program Name         : ����׸� ��� 
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


Dim oPB3S211																	'�� : ����� ComProxy Dll ��� ���� 

Dim StrNextKey											'��: ���� �� 
Dim lgStrPrevKey										'��: ���� �� 
Dim strMode												'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim LngMaxRow											'��: ���� �׸����� �ִ�Row
Dim intFlgMode
Dim LngRow
Dim LngIdx

Dim I1_B_Char
Dim IG1_B_Char_Value

Dim arrCols, arrRows									'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus											'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim	lGrpCnt												'��: Group Count
Dim strCode												'��: Lookup �� ���� ���� 

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

	strMode = Request("txtMode")											'�� : ���� ���¸� ���� 
	intFlgMode = CInt(Request("txtFlgMode"))	
	LngMaxRow = CInt(Request("txtMaxRows"))									'��: �ִ� ������Ʈ�� ���� 
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

	arrRows = Split(Request("txtSpread"), gRowSep)							'��: Spread Sheet ������ ��� �ִ� Element�� 
    ReDim IG1_B_Char_Value(UBound(arrRows,1),2)

	For LngRow = 0 To LngMaxRow - 1
		arrCols = Split(arrRows(LngRow), gColSep)
		strStatus = UCase(Trim(arrCols(0)))									'��: Row �� ���� 

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
       Set oPB3S211 = Nothing		                                        '��: Unload Comproxy DLL
       Response.End		
    End If

	Set oPB3S211 = Nothing                                                  '��: Unload Comproxy

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr																
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr					
	
	'==============================================================================
	' Function : SheetFocus
	' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
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
	' Description : ���ڿ��� ���̸� ���ϴ� ���̿� ���Ͽ� �۰ų� ������ True ���� 
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
