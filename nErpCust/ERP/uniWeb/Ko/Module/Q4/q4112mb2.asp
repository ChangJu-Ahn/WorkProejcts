<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q4112MB2
'*  4. Program Name         : 부적합처리조회 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next		

Call HideStatusWnd
	
Dim PQIG320																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
Dim LngMaxRow
Dim LngRow

Dim arrRowVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim arrColVal								'☜: Spread Sheet 의 값을 받을 Array 변수 
Dim StrStatus								'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)

Dim StrPlantCd
Dim StrInspReqNo
Dim IntInspResultNo
	
Dim ImportData
Dim iErrorPosition

Const Q320_I1_select_char = 0
Const Q320_I1_client_row_num = 1
Const Q320_I1_disposition_cd = 2
Const Q320_I1_qty = 3
Const Q320_I1_remark = 4


LngMaxRow = CLng(Request("txtMaxRows"))					'☜: 최대 업데이트된 갯수 
StrPlantCd = UCase(Request("txtPlantCd"))
StrInspReqNo = UCase(Request("txtInspReqNo"))
intInspResultNo = 1
	
Set PQIG320 = Server.CreateObject("PQIG320.cQMtInspDispSimple")
If CheckSystemError(Err,True) Then											'☜: ComProxy Unload
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

If Request("txtSpread") <> "" Then
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	LngMaxRow = UBound(arrRowVal) - 1
	Redim ImportData(LngMaxRow, 4)
		
	For LngRow = 0 To LngMaxRow
		arrColVal = Split(arrRowVal(LngRow), gColSep)
		StrStatus = UCase(arrColVal(0))
		ImportData(LngRow,Q320_I1_select_char) = StrStatus
		ImportData(LngRow,Q320_I1_disposition_cd) = arrColVal(1)
		If StrStatus = "C" or StrStatus = "U" Then
			ImportData(LngRow,Q320_I1_client_row_num) = arrColVal(4)
			ImportData(LngRow,Q320_I1_Qty) = UniConvNum(arrColVal(2), 0)
			ImportData(LngRow,Q320_I1_Remark) = arrColVal(3)
		Else
			ImportData(LngRow,Q320_I1_client_row_num) = arrColVal(2)
		End if
		
	Next
	
	Call PQIG320.Q_MAINT_INSP_DISPOSIT_SIMPLE_SVR(gStrGlobalCollection, StrPlantCd, StrInspReqNo, intInspResultNo, ImportData, iErrorPosition)
		
	If CheckSYSTEMError(Err,True) Then
		If iErrorPosition > 0 Then
			Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
		End If
		Set PQIG320 = Nothing 
		Response.End
	End if
End If
Set PQIG320 = Nothing                                                   '☜: Unload Comproxy

%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	Dim StrHTML
	If iLoc = I_INSCRIPT Then
		StrHTML = "parent.frm1.vspdData.focus" & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write StrHTML
	ElseIf iLoc = I_MKSCRIPT Then
		
		StrHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.focus" & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		StrHTML = StrHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		StrHTML = StrHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write StrHTML
	End If
End Function
</Script>