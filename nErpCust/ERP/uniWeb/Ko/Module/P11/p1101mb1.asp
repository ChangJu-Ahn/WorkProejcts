<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<%

'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Component List		: PP1G104.cPLkUpLotPeriodSvr.P_LOOK_UP_LOT_PERIOD
'*  7. Modified date(First) : 2000/09/15
'*  8. Modified date(Last)  : 2000/09/26
'*  9. Modifier (First)     : KimTaeHyun
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Dim pPP1G104
Dim I1_prod_work_set_temp_timestamp 
Dim I2_p_mfg_calendar_type_cal_type 
Dim iCommandSent

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
Dim strQryMode
Dim strClnrType
Dim i

Const C_SHEETMAXROWS_D = 100

Dim lgIntPrevKey

Call HideStatusWnd
Call LoadBasisGlobalInf() 

	On Error Resume Next
	Err.Clear
	
	strQryMode = Request("lgIntFlgMode")
    lgIntPrevKey = Request("lgIntPrevKey")
	
	I1_prod_work_set_temp_timestamp  = Trim(Request("txtYear")) & "-01-01"
	I2_p_mfg_calendar_type_cal_type = Trim(Request("txtClnrType"))
	iCommandSent = "LIST"

	Set pPP1G104 = Server.CreateObject("PP1G104.cPLkUpLotPeriodSvr")
	
	If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
	
	call pPP1G104.P_LOOK_UP_LOT_PERIOD (gStrGlobalCollection, I1_prod_work_set_temp_timestamp, _
		I2_p_mfg_calendar_type_cal_type, iCommandSent)
	
	If CheckSYSTEMError(Err, True) = True Then
		Set pPP1G104 = Nothing	
		Response.End
	End If
	
	Set pPP1G104 = Nothing
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	UNISqlId(0) = "180100sab"

	IF Request("txtClnrType") = "" Then
	   strClnrType = "|"
	ELSE
	   strClnrType = FilterVar(Request("txtClnrType"),"''","S")
	END IF

	UNIValue(0, 0) = strClnrType
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
%>
		<Script Language=vbscript>
			parent.frm1.txtClnrTypeNm.value = " "
		</Script>		
<%

	Else
%>
		<Script Language=vbscript>
			parent.frm1.txtClnrTypeNm.value = "<%=ConvSPChars(rs0(0))%>"
		</Script>	
<%
	End If
	
	
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	
	
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)
	
	UNISqlId(0) = "180100saa"
	UNIValue(0, 0) = FilterVar(Request("txtClnrType"),"''","S")
	UNIValue(0, 1) = FilterVar(CInt(Request("txtYear")),"''","S")
	UNIValue(0, 2) = FilterVar(CInt(Request("txtYear")),"''","S")	
	
	Select Case strQryMode
	Case CStr(OPMD_CMODE)
		UNIValue(0, 3) = 0
	Case CStr(OPMD_UMODE)
		UNIValue(0, 3) = UCase(Trim(lgIntPrevKey))		
	End Select
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox(900014, vbOKOnly, "", "", I_MKSCRIPT)		
		rs0.Close
		Set rs0 = Nothing
					
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent
	LngMaxRow = .frm1.vspdData.MaxRows
		
<%  
	If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
		ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
	Else
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
	End If
	
    For i=0 to rs0.RecordCount-1 
		If I < C_SHEETMAXROWS_D Then
%>  
			strData = ""                                                            
			strData = strData & Chr(11) & "<%=rs0("LOT_PERD_NO")%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("START_DT"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("END_DT"))%>"
			strData = strData & Chr(11) & "<%=rs0("OPR_DAY_WITHIN_PERD")%>"
			strData = strData & Chr(11) & "<%=i%>"
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		End If
	Next
%>
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowDataByClip iTotalStr

	.lgIntPrevKey = "<%=rs0("LOT_PERD_NO")%>"
		
	If .lgIntPrevKey = "" Then
		.lgIntPrevKey = 0		
	End If
		
	.frm1.hClnrType.value = "<%=ConvSPChars(Request("txtClnrType"))%>"
	.frm1.hYear.value = "<%=Request("txtYear")%>"
		
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbQueryOk

End With	
</Script>	
<%
Set ADF = Nothing
%>

<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
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
</Script>
