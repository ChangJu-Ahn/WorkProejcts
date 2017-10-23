<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1504mb2.asp
'*  4. Program Name         : �ڿ��� Shift Save
'*  5. Program Desc         :
'*  6. Comproxy List        : + 
'*  7. Modified date(First) : 2000/09/20
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : Ryu Sung Won
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%
Call LoadBasisGlobalInf


Dim oPP1G612											'�� : ����� ComProxy Dll ��� ���� 

Dim strMode												'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim intFlgMode
Dim lgStrPrevKey										'��: ���� �� 
Dim LngMaxRow											'��: ���� �׸����� �ִ�Row
Dim LngRow

Dim arrRows, arrCols									'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus											'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim strCode												'��: Lookup �� ���� ���� 
Dim ChkTimeVal1											'��: �ð�data�� ����� �Է���� üũ�ϴ� ���� 
Dim ChkTimeVal2
	
Dim I1_P_Shift_Header_Shift_Cd
Dim I2_P_Resource_Resource_Cd
Dim I3_B_Plant_Plant_Cd
Dim IG1_Import_Group

Const C_IG1_I1_select_char = 0
Const C_IG1_I2_shift_exception_cd = 1
Const C_IG1_I2_description = 2
Const C_IG1_I2_exception_type = 3
Const C_IG1_I2_start_dt = 4
Const C_IG1_I2_end_dt = 5
Const C_IG1_I2_work_flg = 6
	
Call HideStatusWnd

On Error Resume Next
Err.Clear
	
	strMode = Request("txtMode")
	
	LngMaxRow    = CInt(Request("txtMaxRows"))
	intFlgMode   = CInt(Request("txtFlgMode"))
    lgStrPrevKey = Trim(Request("lgStrPrevKey"))
    
    If intFlgMode = OPMD_CMODE Then
		I3_B_Plant_Plant_Cd			= UCase(Request("txtPlantCd"))
		I2_P_Resource_Resource_Cd	= UCase(Request("txtResourceCd"))
		I1_P_Shift_Header_Shift_Cd	= UCase(Request("txtShiftCd"))		
    Else
		I3_B_Plant_Plant_Cd			= UCase(Request("hPlantCd"))
		I2_P_Resource_Resource_Cd	= UCase(Request("hResourceCd"))
		I1_P_Shift_Header_Shift_Cd	= UCase(Request("hShiftCd"))
    End If

    If I3_B_Plant_Plant_Cd = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If
	
	If I2_P_Resource_Resource_Cd = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If
	If I1_P_Shift_Header_Shift_Cd = "" Then
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End If
		

	arrRows = Split(Request("txtSpread"), gRowSep)							'��: Spread Sheet ������ ��� �ִ� Element�� 
    ReDim IG1_Import_Group(UBound(arrRows,1),6)

	For LngRow = 0 To LngMaxRow - 1
    
		arrCols = Split(arrRows(LngRow), gColSep)
		strStatus = UCase(Trim(arrCols(0)))									'��: Row �� ���� 

		Select Case strStatus
			
		    Case "C", "U"
				IG1_Import_Group(LngRow, C_IG1_I1_select_char)			= strStatus
				IG1_Import_Group(LngRow, C_IG1_I2_shift_exception_cd)	= UCase(arrCols(2))
				IG1_Import_Group(LngRow, C_IG1_I2_description)			= arrCols(3)
				IG1_Import_Group(LngRow, C_IG1_I2_work_flg)				= arrCols(4)
				IG1_Import_Group(LngRow, C_IG1_I2_exception_type)		= arrCols(9)
				
				ChkTimeVal1 = ConvToSec(arrCols(6))
				ChkTimeVal2 = ConvToSec(arrCols(8)) 

				If UniConvDate(arrCols(5)) > UniConvDate(arrCols(7)) Then
					Call DisplayMsgBox("972002", vbInformation, "������","������", I_MKSCRIPT)
					Call SheetFocus(arrCols(1),4,I_MKSCRIPT)
					Response.End	
				End If
				
				If ChkTimeVal1 > ChkTimeVal2 Then
					Call DisplayMsgBox("972002", vbInformation, "����ð�","���۽ð�", I_MKSCRIPT)
					Call SheetFocus(arrCols(1),5,I_MKSCRIPT)
					Response.End	
				End If
				
				If Len(Trim(arrCols(5))) Then
					If UniConvDate(arrCols(5)) = "" Then	 
						Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
						Call SheetFocus(arrCols(1),5,I_MKSCRIPT)
						Response.End	
					ElseIf ChkTimeVal1 = -999999	Then
						Call DisplayMsgBox("970029", vbInformation, "���۽ð�", "", I_MKSCRIPT)
						Call SheetFocus(arrCols(1),5,I_MKSCRIPT)
						Response.End
					Else	
						IG1_Import_Group(LngRow, C_IG1_I2_start_dt) = UniConvDate(arrCols(5)) & " " & arrCols(6)
					End If
				End If
				
				If Len(Trim(arrCols(7))) Then
					If UniConvDate(arrCols(7)) = "" Then	 
						Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
						Call SheetFocus(arrCols(1),7,I_MKSCRIPT)
						Response.End	
					ElseIf ChkTimeVal2 = -999999	Then
						Call DisplayMsgBox("970029", vbInformation, "����ð�", "", I_MKSCRIPT)
						Call SheetFocus(arrCols(1),7,I_MKSCRIPT)
						Response.End
					Else	
						IG1_Import_Group(LngRow, C_IG1_I2_end_dt) = UniConvDate(arrCols(7)) & " " & arrCols(8)
					End If
				End If
	
			Case "D"
				IG1_Import_Group(LngRow, C_IG1_I1_select_char) = "D"
				IG1_Import_Group(LngRow, C_IG1_I2_shift_exception_cd) = UCase(arrCols(2))
	
		End Select
        
        If LngRow >= 99 Or LngRow = LngMaxRow - 1 Then							'��: 5���� Group����, ������ �϶� 
			Exit For
		End If
	
	Next


	Set oPP1G612 = Server.CreateObject("PP1G612.cPMngShiftExcpt")    

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End 
	End If	

	Call oPP1G612.P_MANAGE_SHIFT_EXCEPTION(gStrGlobalCollection, _
									I1_P_Shift_Header_Shift_Cd, _
									I2_P_Resource_Resource_Cd, _
									I3_B_Plant_Plant_Cd, _
									IG1_Import_Group)	
		
	If CheckSYSTEMError(Err,True) = True Then
       Set oPP1G612 = Nothing
       Response.End		
    End If
    
    Set oPP1G612 = Nothing
    
    Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr
	Response.Write "		.lgStrPrevKey = """ & ConvSPChars(lgStrPrevKey) & """" & vbCr
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr



'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================

'==============================================================================
' Function : ConvToSec()
' Description : ����ÿ� �� �ð� �����͵��� �ʷ� ȯ�� 
'==============================================================================
Function ConvToSec(ByVal Str)
	
	If Str = "" Then
		ConvToSec = 0
	ElseIf Len(Trim(Str)) = 8 Then
		ConvToSec = CInt(Trim(Mid(Str,1,2))) * 3600 + CInt(Trim(Mid(Str,4,2))) * 60 + CInt(Trim(Mid(Str,7,2)))
	Else
		ConvToSec = -999999
	End If

End Function

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
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
%>