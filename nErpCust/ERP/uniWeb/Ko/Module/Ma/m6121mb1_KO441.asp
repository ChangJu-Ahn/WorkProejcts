<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M6121MB1
'*  4. Program Name         : 부대비일괄배부취소 
'*  5. Program Desc         : 부대비일괄배부취소 
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/11/05
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 

On Error Resume Next  

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2					'DBAgent Parameter 선언 
Dim strQryMode								'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Const C_SHEETMAXROWS_D = 100

Dim lgOpModeCRUD

Err.Clear 
                                             '☜: Clear Error status
Call HideStatusWnd                                                               '☜: Hide Processing message
	
lgOpModeCRUD  = Request("txtMode") 
								                                              '☜: Read Operation Mode (CRUD)
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                      '☜: Query
         Call  SubBizQueryMulti()
    Case CStr(UID_M0002)                                                      '☜: Save,Update
         Call SubBizSave()                                                       '☜: Delete
End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()	
	Dim strPlantCd
	Dim strProcessStep
	Dim strDisbFrDt
	Dim strDisbToDt
	Dim strBatchJobFrDt
	Dim strBatchJobToDt
	Dim strDocumentNo
	
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr
	Dim i
	Dim LoopCnt
	Dim strNextKey
	Dim LngMaxRow
	
	strPlantCd = Ucase(Trim(Request("txtPlantCd")))
	strProcessStep = Ucase(Trim(Request("txtProcessStep")))
	strDisbFrDt = Ucase(Trim(Request("txtFrDisbDt")))
	strDisbToDt = Ucase(Trim(Request("txtToDisbDt")))
	strBatchJobFrDt = Ucase(Trim(Request("txtFrBatchJobDt")))
	strBatchJobToDt = Ucase(Trim(Request("txtToBatchJobDt")))
	strDocumentNo = FilterVar(Ucase(Trim(Request("lgStrPrevKey"))),"''","S")
	
	LngMaxRow = Clng(Request("txtMaxRows"))
	
	If strPlantCd = "" then
		strPlantCd = "|"
	else
		strPlantCd = FilterVar(strPlantCd,"''","S")
	end if
	
	If strProcessStep = "" then
		strProcessStep = "|"
	else
		strProcessStep = FilterVar(strProcessStep,"''","S")
	end if
	
	If strDisbFrDt = "" then
		strDisbFrDt = "|"
	else
		strDisbFrDt = FilterVar(UNIConvDate(strDisbFrDt),"''","S")
	end if
	
	If strDisbToDt = "" then
		strDisbToDt = "|"
	else
		strDisbToDt = FilterVar(UNIConvDate(strDisbToDt),"''","S")
	end if
	
	If strBatchJobFrDt = "" then
		strBatchJobFrDt = "|"
	else
		strBatchJobFrDt = FilterVar(UNIConvDate(strBatchJobFrDt),"''","S")
	end if
	
	If strBatchJobToDt = "" then
		strBatchJobToDt = "|"
	else
		strBatchJobToDt = FilterVar(UNIConvDate(strBatchJobToDt),"''","S")
	end if
	
	Redim UNISqlId(2)
	Redim UNIValue(2, 6)

	UNISqlId(0) = "M6121MA101"
	UNISqlId(1) = "M2111QA302"
	UNISqlId(2) = "M6111QA103"

	UNIValue(0, 0) = strPlantCd
	UNIValue(0, 1) = strProcessStep
	UNIValue(0, 2) = strDisbFrDt
	UNIValue(0, 3) = strDisbToDt
	UNIValue(0, 4) = strBatchJobFrDt
	UNIValue(0, 5) = strBatchJobToDt
	UNIValue(0, 6) = strDocumentNo
	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(2, 0) = FilterVar(Ucase(Trim(Request("txtProcessStep"))),"''","S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
		
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	if Request("lgIntFlgMode") = CStr(OPMD_CMODE) then
		Redim UNISqlId(2)
		Redim UNIValue(2, 6)

		UNISqlId(0) = "M6121MA101"
		UNISqlId(1) = "M2111QA302"
		UNISqlId(2) = "M6111QA103"

		UNIValue(0, 0) = strPlantCd
		UNIValue(0, 1) = strProcessStep
		UNIValue(0, 2) = strDisbFrDt
		UNIValue(0, 3) = strDisbToDt
		UNIValue(0, 4) = strBatchJobFrDt
		UNIValue(0, 5) = strBatchJobToDt
		UNIValue(0, 6) = strDocumentNo
		UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
		UNIValue(2, 0) = FilterVar(Ucase(Trim(Request("txtProcessStep"))),"''","S")
				
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
		
		if strPlantCd <> "|" then
			If (rs1.EOF And rs1.BOF) Then
				rs1.Close
				Set rs1 = Nothing
				Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
				Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
				Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf	
				Response.Write "</Script>" & vbCrLf
			Else
				Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				rs1.Close
				Set rs1 = Nothing
			End If
		end if
		
		if strProcessStep <> "|" then
			If (rs2.EOF And rs2.BOF) Then
				rs2.Close
				Set rs2 = Nothing
				Call DisplayMsgBox("176131", vbOKOnly, "", "", I_MKSCRIPT)
				Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtProcessStepNm.value = """"" & vbCrLf
				Response.Write "parent.frm1.txtProcessStep.focus" & vbCrLf	
				Response.Write "</Script>" & vbCrLf
			Else
				Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtProcessStepNm.value = """ & ConvSPChars(rs2("MINOR_NM")) & """" & vbCrLf	'@@rs1("PLANT_NM")수정[060524]
				Response.Write "</Script>" & vbCrLf
				rs2.Close
				Set rs2 = Nothing
			End If
		end if	
	else
		Redim UNISqlId(0)
		Redim UNIValue(0, 6)
		
		UNIValue(0, 0) = strPlantCd
		UNIValue(0, 1) = strProcessStep
		UNIValue(0, 2) = strDisbFrDt
		UNIValue(0, 3) = strDisbToDt
		UNIValue(0, 4) = strBatchJobFrDt
		UNIValue(0, 5) = strBatchJobToDt
		UNIValue(0, 6) = strDocumentNo
		UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
		UNIValue(2, 0) = FilterVar(Ucase(Trim(Request("txtProcessStep"))),"''","S")
		
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	end if
	
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	else
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
			LoopCnt = C_SHEETMAXROWS_D - 1
		Else
			LoopCnt = rs0.RecordCount - 1
		End If
		
		ReDim TmpBuffer(LoopCnt)
		
		For i=0 to LoopCnt
			strData = ""
			strData = strData & Chr(11) & ConvSPChars(rs0("Item_Document_No"))
			strData = strData & Chr(11) & ConvSPChars(rs0("dist_ref_no"))
			strData = strData & Chr(11) & ConvSPChars(rs0("Plant_Cd"))
			strData = strData & Chr(11) & ConvSPChars(rs0("Plant_Nm"))
			strData = strData & Chr(11) & UNIDateClientFormat(ConvSPChars(rs0("Disb_Dt")))
			strData = strData & Chr(11) & UNIDateClientFormat(ConvSPChars(rs0("Disb_Job_Dt")))
			strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("tot_disb_amt"),0)
			strData = strData & Chr(11) & UNIDateClientFormat(ConvSPChars(rs0("Disb_Qry_Fr_Dt")))
			strData = strData & Chr(11) & UNIDateClientFormat(ConvSPChars(rs0("Disb_Qry_To_Dt")))
			strData = strData & Chr(11) & ConvSPChars(rs0("Process_step"))
			strData = strData & Chr(11) & ConvSPChars(rs0("minor_Nm"))
			strData = strData & Chr(11) & LngMaxRow + i
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(i) = strData
			rs0.MoveNext
		Next
		
		if Not (rs0.eof) then
			strNextKey = rs0("Item_Document_No")
		end if
		
		iTotalStr = Join(TmpBuffer, "")
		
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
	End If
	
	'Call ServerMesgBox(iTotalStr, vbCritical, I_MKSCRIPT) 
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "	.ggoSpread.Source          =  .frm1.vspdData " & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & iTotalStr & """" & vbCr	
    Response.Write "	.frm1.vspdData.Redraw = false " & vbCr
    Response.Write "	.frm1.vspdData.Redraw = True " & vbCr
    Response.Write "	.lgStrPrevKey              = """ & StrNextKey & """" & vbCr  
    Response.Write " .frm1.hPlantCd.value     = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCr
	Response.Write " .frm1.hProcessStep.value   = """ & ConvSPChars(Request("txtProcessStep"))   & """" & vbCr
	'Response.Write " .frm1.hFrDisbDt.value    = """ & UNIDateClientFormat(Request("txtFrDisbDt")) & """" & vbCr   'KSJ 수정 
	'Response.Write " .frm1.hToDisbDt.value     = """ & UNIDateClientFormat(Request("txtToDisbDt")) & """" & vbCr  'KSJ 수정 
	Response.Write " .frm1.hFrDisbDt.value    = """ & Request("txtFrDisbDt") & """" & vbCr
	Response.Write " .frm1.hToDisbDt.value     = """ & Request("txtToDisbDt") & """" & vbCr
	'Response.Write " .frm1.hFrBatchJobDt.value   = """ & UNIDateClientFormat(Request("txtFrBatchJobDt")) & """" & vbCr  'KSJ 수정 
	'Response.Write " .frm1.hToBatchJobDt.value   = """ & UNIDateClientFormat(Request("txtToBatchJobDt")) & """" & vbCr  'KSJ 수정 
	Response.Write " .frm1.hFrBatchJobDt.value   = """ & Request("txtFrBatchJobDt") & """" & vbCr
	Response.Write " .frm1.hToBatchJobDt.value   = """ & Request("txtToBatchJobDt") & """" & vbCr
    Response.Write " .DbQueryOk "		    	  & vbCr 
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr  
	
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
	On Error Resume Next
    Err.Clear	
    
    Dim strDocumentNo
    Dim strRefNo
    Dim strDisbBatchJobDt
    Dim strProcessStep
    Dim strPlantCd
	Dim strDisbQryDt
	Dim strDisbFrQryDt
	
	Dim iPMAG182

	strDocumentNo = Ucase(Trim(Request("hdnDocumentNo")))
	strProcessStep = Ucase(Trim(Request("hdnProcessStep")))	'@@ hdnProcesssStep -> hdnProcessStep 변경[060524]
	strPlantCd = Ucase(Trim(Request("hdnPlantCd")))
	strDisbQryDt = Ucase(Trim(Request("hdnDisbQryDt")))
	strDisbFrQryDt = Ucase(Trim(Request("hdnDisbFrQryDt")))
	'strDisbBatchJobDt = Ucase(Trim(Request("hdnDisbBatchJobDt")))
	strDisbBatchJobDt = UNIConvDate(Ucase(Trim(Request("hdnDisbBatchJobDt"))))  'KSJ 수정 
	strRefNo = UCase(Trim(Request("hdnRefNo")))
    
    Set iPMAG182 = Server.CreateObject("PMAG182.cMMaintDistSvr")    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
		Set iPMAG182 = Nothing 		
		Exit Sub
	End If
    
    Call iPMAG182.M_MAINT_DISTRIBUT_SVR(gStrGlobalCollection, "C", strPlantCd, strProcessStep, strDisbQryDt, strDisbBatchJobDt, _
										strRefNo,	strDocumentNo, strDisbFrQryDt)
    
    If CheckSYSTEMError(Err,True) = true then
		Set iPMAG182 = Nothing 		
		Exit Sub
	End If
    
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "   
    																	'☜: Protect system from crashing        
End Sub    


'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	
	If Trim(lRow) = "" Then Exit Function
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
