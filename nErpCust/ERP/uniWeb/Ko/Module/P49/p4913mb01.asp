<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4913mb01.asp
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005-01-27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd

Dim pPP4G903

Dim lgIntFlgMode
Dim LngMaxRow

Dim arrRowVal								'��: Spread Sheet �� ���� ���� Array ���� 
Dim arrColVal								'��: Spread Sheet �� ���� ���� Array ���� 

Dim iErrorPosition
Dim iErrorSpread
Dim LngRow
Dim txtSpread
Dim txtSpread1
Dim txtSpread2
Dim txtSpread3
Dim txtSpread4
Dim txtSpread5

Dim strPlantCd
Dim strReportDt
Dim strWcCd

strPlantCd = UCase(Trim(Request("txtPlantCd")))
strReportDt = Trim(Request("txtprodDt"))
strWcCd = UCase(Trim(Request("txtWcCd")))

txtSpread = Request("txtSpread")		' Create, Update
txtSpread1 = Request("txtSpread1")		' Create, Update
txtSpread2 = Request("txtSpread2")		' Create, Update
txtSpread3 = Request("txtSpread3")		' Create, Update
txtSpread4 = Request("txtSpread4")		' Create, Update
txtSpread5 = Request("txtSpread5")		' Create, Update

'-------------------------------------------------------------------------------
'	COM+ Action
'-------------------------------------------------------------------------------

Set pPP4G903 = Server.CreateObject("PP4G903.cPMngPrdtDailyReport")
                                            
If CheckSYSTEMError(Err,True) = True Then
   Response.End
End If

Call pPP4G903.P_MANAGE_PRDT_DAILY_REPORT(gStrGlobalCollection, strPlantCd, strReportDt, strWcCd, txtSpread, txtSpread1, txtSpread2, txtSpread3, txtSpread4, txtSpread5, iErrorPosition, iErrorSpread)

'-------------------------------------------------------------------------------
'	GRID ��Ŀ�� ó�� 
'-------------------------------------------------------------------------------
If CheckSYSTEMError2(Err,True, "","","","","") = True Then
	If iErrorPosition <> "" Then
		'Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT, iErrorSpread)
		Set pPP4G903 = Nothing
		Response.End
	End If
End If
'-------------------------------------------------------------------------------

Set pPP4G903 = Nothing

%>
<Script Language=vbscript>

Select Case parent.gSelframeFlg
	Case 1	'TAB1
		parent.DbSaveOk																		'��: ȭ�� ó�� ASP �� ��Ī�� 
	Case 2	'TAB2
		parent.DbSaveFormOk																	'��: ȭ�� ó�� ASP �� ��Ī�� 
End Select

</Script>
<%
' Server Side ������ ���⼭ ���� 
'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc, ByVal iSpread)
	Dim strHTML
	If iLoc = I_INSCRIPT Then
	
		Select Case iSpread
			Case "txtSpread1"
				strHTML = "With parent.frm1.vspdData1" & vbCrLf
			Case "txtSpread2"
				strHTML = "With parent.frm1.vspdData2" & vbCrLf
			Case "txtSpread3"
				strHTML = "With parent.frm1.vspdData3" & vbCrLf
			Case "txtSpread4"
				strHTML = "With parent.frm1.vspdData4" & vbCrLf
			Case "txtSpread5"
				strHTML = "With parent.frm1.vspdData5" & vbCrLf
		End Select 

		strHTML = strHTML & "	.focus" & vbCrLf
		strHTML = strHTML & "	.Row = " & lRow & vbCrLf
		strHTML = strHTML & "	.Col = " & lCol & vbCrLf
		strHTML = strHTML & "	.Action = 0" & vbCrLf
		strHTML = strHTML & "	.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "	.SelLength = len(.Text) " & vbCrLf
		strHTML = strHTML & "End With " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf

		Select Case iSpread
			Case "txtSpread1"
				strHTML = strHTML & "With parent.frm1.vspdData1" & vbCrLf
			Case "txtSpread2"
				strHTML = strHTML & "With parent.frm1.vspdData2" & vbCrLf
			Case "txtSpread3"
				strHTML = strHTML & "With parent.frm1.vspdData3" & vbCrLf
			Case "txtSpread4"
				strHTML = strHTML & "With parent.frm1.vspdData4" & vbCrLf
			Case "txtSpread5"
				strHTML = strHTML & "With parent.frm1.vspdData5" & vbCrLf
		End Select 

		strHTML = strHTML & "	.focus" & vbCrLf
		strHTML = strHTML & "	.Row = " & lRow & vbCrLf
		strHTML = strHTML & "	.Col = " & lCol & vbCrLf
		strHTML = strHTML & "	.Action = 0" & vbCrLf
		strHTML = strHTML & "	.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "	.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "End With " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
</Script>