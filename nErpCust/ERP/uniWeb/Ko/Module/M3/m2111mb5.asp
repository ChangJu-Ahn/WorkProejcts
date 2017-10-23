<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m2111mb5
'*  4. Program Name         : 구매요청조정등록 
'*  5. Program Desc         : 구매요청조정등록 
'*  6. Component List       : PM2G151.cMAmendPR / PM2G139.cMLookupSpplLtS
'*  7. Modified date(First) : 2002/07/02
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrPoNo
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim iLngMaxRow		' 현재 그리드의 최대Row
Dim iLngRow
Dim GroupCount  
Dim lgCurrency        
Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수     
Dim lgDataExist
Dim lgPageNo
Dim PlantNm			'☜ : 공장명 
Dim ItemNm			'☜ : 품목명 
Dim PrTypeNm		'☜ : 요청구분명 
Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows
intARows=0
intTRows=0
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message

lgOpModeCRUD  = Request("txtMode") 

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
		 Call  SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)
         Call SubBizSaveMulti()
	Case "LookUpItemPlant"
		 Call SubLookUpItemPlant()
    Case "LookSppl"				'☜: 공급처 Change Event
		 Call SubLookSppl
End Select

Sub SubBizQueryMulti()
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")
	
	Call FixUNISQLData()		'☜ : DB-Agent로 보낼 parameter 데이타 set
	
	Call QueryData()			'☜ : DB-Agent를 통한 ADO query
	
	'-----------------------
	'Result data display area
	'----------------------- 
%>
	<Script Language=vbscript>
		With parent
			.frm1.hdnPlant.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
			
			.frm1.txtPlantNm.Value	= "<%=PlantNm%>"
			.frm1.txtItemNm.Value	= "<%=ConvSPChars(ItemNm)%>"
			.frm1.txtPrTypeNm.Value	= "<%=PrTypeNm%>"						
			
			.frm1.hdnItem.value = "<%=ConvSPChars(Request("txtitemcd"))%>"
			.frm1.hdnRFrDt.value = "<%=ConvSPChars(Request("txtReqFrDt"))%>"
			.frm1.hdnRToDt.value = "<%=ConvSPChars(Request("txtReqToDt"))%>"
			.frm1.hdnDFrDt.value = "<%=ConvSPChars(Request("txtDlvyFrDt"))%>"
			.frm1.hdnDToDt.value = "<%=ConvSPChars(Request("txtDlvyToDt"))%>"

			.frm1.hdnPrTypeCd.value = "<%=ConvSPChars(Request("txtPrTypeCd"))%>"			
			.frm1.hdnMrp.value = "<%=ConvSPChars(Request("txtMRP"))%>"
			.frm1.hdnTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
			
			.frm1.txtPlantCd.focus
			Set .gActiveElement = .document.activeElement

			If "<%=lgDataExist%>" = "Yes" Then
				
				'Show multi spreadsheet data from this line
				       
				.ggoSpread.Source    = .frm1.vspdData 
				.ggoSpread.SSShowData "<%=istrData%>"                  '☜: Display data 
				
				.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
				
				.DbQueryOk <%=intARows%>,<%=intTRows%>
							
			End If
		End with
	</Script>	
<%	
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	Dim M21115
	dim iLngMaxRow
	Dim lgIntFlgMode
	Dim iStrCommandSent
	Dim iErrorPosition
	Dim iStrSpread
	Dim LngRow
	Dim arrValUp, arrValDn
	Dim arrTemp1, arrTemp2
	Dim lgTransSep
	Dim lgHdDtlSep
	Dim iInti 
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim ii
    
    Dim arrForRowNo
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next

    itxtSpread = Join(itxtSpreadArr,"")

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   

    lgTransSep = "º"
    lgHdDtlSep = "Ð"

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear		
	
	arrTemp1	= Split(itxtSpread, lgTransSep)		
	
	If ubound(arrTemp1,1) > 0 Then

		Set M21115 = Server.CreateObject("PM2G151.cMAmendPR") 
	
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		    
		If CheckSYSTEMError(Err,True) = true Then 		
			Set M21115 = Nothing												'☜: ComPlus Unload
			Exit Sub														'☜: 비지니스 로직 처리를 종료함 
		End if
	
		For iInti=0 to ubound(arrTemp1)-1
			arrTemp2 = Split(arrTemp1(iInti), lgHdDtlSep)
			arrValUp = arrTemp2(0)
			arrValDn = arrTemp2(1)

			Call M21115.M_AMEND_PR(gStrGlobalCollection, arrValUp, arrValDn, iErrorPosition)
			
			If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then 			
				Set M21115 = Nothing                  							
				exit sub	
			Else 			
				arrForRowNo = split(arrValUp,gColSep)
				'처리가 완료된것은 Check Box 가 풀림.
				Response.Write "<Script language=vbscript> " & vbCr         
				Response.Write "	with Parent.frm1.vspdData"      & vbCr	
				Response.Write "		.Col = parent.C_CfmFlg	"      & vbCr		
				Response.Write "		.Row = " & arrForRowNo(ubound(arrForRowNo))   & vbCr					
				Response.Write "		.Text = 0 "   & vbCr		
				Response.Write "	end with "      & vbCr	
				Response.Write "</Script> "             & vbCr	
			End If
			
			arrValUp = ""
			arrValDn = ""
		Next
	End IF
	
	If NOT(M21115 is Nothing) Then
		Set M21115 = Nothing                                                   '☜: Unload Comproxy
	End If	  
	
	Response.Write "<Script language=vbs> " & vbCr
	Response.Write "       Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "</Script> "   & vbCr
	       
End Sub    

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim PvArr
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt

	Const M_PUR_REQ_PR_NO				= 0
	Const M_PUR_REQ_PROCURE_TYPE		= 1
	Const M_PUR_REQ_PR_TYPE				= 2
	Const M_PUR_REQ_PR_STS				= 3
	Const M_PUR_REQ_REQ_QTY				= 4
	Const M_PUR_REQ_REQ_UNIT			= 5
	Const M_PUR_REQ_REQ_CFM_QTY			= 6
	Const M_PUR_REQ_BASE_REQ_QTY		= 7
	Const M_PUR_REQ_BASE_REQ_UNIT		= 8
	Const M_PUR_REQ_ORD_QTY				= 9
	Const M_PUR_REQ_RCPT_QTY			= 10
	Const M_PUR_REQ_IV_QTY				= 11
	Const M_PUR_REQ_REQ_DT				= 12
	Const M_PUR_REQ_DLVY_DT				= 13
	Const M_PUR_REQ_PUR_PLAN_DT			= 14
	Const M_PUR_REQ_REQ_DEPT			= 15
	Const M_PUR_REQ_REQ_PRSN			= 16
	Const M_PUR_REQ_SPPL_CD				= 17
	Const M_PUR_REQ_SL_CD				= 18
	Const M_PUR_REQ_PUR_ORG				= 19
	Const M_PUR_REQ_PUR_GRP				= 20
	Const M_PUR_REQ_MRP_ORD_NO			= 21
	Const M_PUR_REQ_MRP_RUN_NO			= 22
	Const M_PUR_REQ_TRACKING_NO			= 23
	Const M_PUR_REQ_SO_NO				= 24
	Const M_PUR_REQ_SO_SEQ_NO			= 25
	Const M_PUR_REQ_plant_cd			= 26
	Const M_PUR_REQ_item_cd				= 27
	Const B_PLANT_PLANT_NM				= 28
	Const B_ITEM_ITEM_NM				= 29
	Const B_ITEM_SPEC					= 30
	Const b_biz_partner_bp_nm			= 31
	Const b_pur_org_pur_org_nm			= 32
	Const b_pur_grp_pur_grp_nm			= 33
	Const b_minor_pr_sts_nm				= 34
	Const b_minor_pr_type_nm			= 35
    
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       intTRows		= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
	
	'----- 레코드셋 칼럼 순서 ----------
	'a.PR_NO, a.PR_TYPE, a.PR_STS, a.PLANT_CD, a.ITEM_CD, a.REQ_QTY, a.REQ_UNIT, a.REQ_CFM_QTY, a.BASE_REQ_QTY, a.BASE_REQ_UNIT, 
	'a.ORD_QTY, a.RCPT_QTY, a.IV_QTY, a.REQ_DT, a.DLVY_DT, a.PUR_PLAN_DT, a.REQ_DEPT, a.REQ_PRSN, a.MRP_ORD_NO, a.MRP_RUN_NO, 
	'b.PLANT_NM, c.ITEM_NM, c.SPEC, d.minor_nm pr_sts_nm, e.minor_nm pr_type_nm 
	'-----------------------------------
	iLoopCount = 0
    ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		
		iRowStr = ""
		iRowStr = iRowStr & Chr(11) & "0"		 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_plant_cd))		    '2공장		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_item_cd))			'3품목 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM_ITEM_NM))		    '4품목명 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM_SPEC))		    '5품목규격 
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_PUR_REQ_REQ_QTY),ggExchRate.DecPoint,0)			'6요청량 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_REQ_UNIT))	        '7단위   
		iRowStr = iRowStr & Chr(11) & ""							'8단위팝업 
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0(M_PUR_REQ_DLVY_DT))	'9필요일 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_PUR_ORG))			'10 구매조직 
		iRowStr = iRowStr & Chr(11) & ""											'11 구매조직 팝업			
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(b_pur_org_pur_org_nm))	    '12 구매조직명 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_PR_NO))             '13요청번호 
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0(M_PUR_REQ_REQ_DT))	'14요청일 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_PR_STS))	        '15요청상태 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(b_minor_pr_sts_nm))	        '16요청상태명 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_PR_TYPE))	        '17요청구분 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(b_minor_pr_type_nm))	        '18요청구분명 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_MRP_RUN_NO))	    '19MRP Run 번호 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_REQ_DEPT))	        '20요청부서 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_REQ_PRSN))	        '21요청자 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_REQ_TRACKING_NO))	        '22 Tracking_No 200308
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             
		
			      
		If iLoopCount - 1 < C_SHEETMAXROWS_D Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount-1) = istrData	
		   istrData = ""
		Else
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If
		
		rs0.MoveNext
	Loop
	istrData = Join(PvArr, "")
	
		
	intARows = iLoopCount
	If iLoopCount =< C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
	  lgPageNo = ""
	End If
		    
	rs0.Close                                                       '☜: Close recordset object
	Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    SetConditionData = false
    
	If Not(rs1.EOF Or rs1.BOF) Then
		PlantNm = rs1("Plant_Nm")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtPlantCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    exit function
		End If
	End If   	
	
	If Not(rs2.EOF Or rs2.BOF) Then
		ItemNm = rs2("Item_Nm")
		Set rs2 = Nothing
	Else
		Set rs2 = Nothing
		If Len(Request("txtitemcd")) Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    rs0.Close
		    Set rs0 = Nothing
		    Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtItemCd.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
		    exit function
		End If
	End If   	
	
	If Not(rs3.EOF Or rs3.BOF) Then
		PrTypeNm = rs3("Minor_Nm")
		Set rs3 = Nothing
	Else
		Set rs3 = Nothing
		If Len(Request("txtPrTypeCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "요청구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    exit function
		End If
	End If   	

    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	dim sTemp
	Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,9)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
	strVal = ""
    UNISqlId(0) = "M2111MA501"
    UNISqlId(1) = "M2111QA302"		'공장조회 
    UNISqlId(2) = "M2111QA303"		'품목조회 
    UNISqlId(3) = "M2111QA306"		'요청구분조회 
    
	UNIValue(1,0) = "" & FilterVar("zzzzz", "''", "S") & ""
    UNIValue(2,0) = "" & FilterVar("zzzzz", "''", "S") & ""
    UNIValue(2,1) = "" & FilterVar("zzzzzzzzzz", "''", "S") & ""
    UNIValue(3,0) = "" & FilterVar("zzzz", "''", "S") & ""
    
    sTemp = "1"
    
    '공장 
    UNIValue(0,0) = "^" 
    If Trim(Request("txtPlantCd")) <> "" Then
		UNIValue(0,1) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	    UNIValue(1,0) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	Else 
	    UNIValue(0,1) = "|"
	End If
	
    '품목 
    If Trim(Request("txtitemcd")) <> "" Then
		UNIValue(0,2) = "  " & FilterVar(Trim(UCase(Request("txtitemcd"))), " " , "S") & "  "
	    UNIValue(2,0) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	    UNIValue(2,1) = "  " & FilterVar(Trim(UCase(Request("txtitemcd"))), " " , "S") & "  "
	Else 
	    UNIValue(0,2) = "|"
	End If
	
    '요청일 
    If Trim(Request("txtReqFrDt")) <> "" Then
		UNIValue(0,3) =  "  " & FilterVar(UniConvDate(Request("txtReqFrDt")), "''", "S") & " "	
    Else
        UNIValue(0,3) = "|"
	End If
			
    If Trim(Request("txtReqToDt")) <> "" Then
		UNIValue(0,4) =  "  " & FilterVar(UniConvDate(Request("txtReqToDt")), "''", "S") & " "	
    Else
        UNIValue(0,4) = "|"
	End If
	
    '필요일 
    If Trim(Request("txtDlvyFrDt")) <> "" Then
		UNIValue(0,5) =  "  " & FilterVar(UniConvDate(Request("txtDlvyFrDt")), "''", "S") & " "	
    Else
        UNIValue(0,5) = "|"
	End If
			
    If Trim(Request("txtDlvyToDt")) <> "" Then
		UNIValue(0,6) =  "  " & FilterVar(UniConvDate(Request("txtDlvyToDt")), "''", "S") & " "	
    Else
        UNIValue(0,6) = "|"
	End If
	
    '요청구분 
    If Trim(Request("txtPrTypeCd")) <> "" Then
		UNIValue(0,7) = "  " & FilterVar(Trim(UCase(Request("txtPrTypeCd"))), " " , "S") & "  "
	    UNIValue(3,0) = "  " & FilterVar(Trim(UCase(Request("txtPrTypeCd"))), " " , "S") & "  "
	Else 
	    UNIValue(0,7) = "|"
	End If
	
    'MRP Run 번호 
    If Trim(Request("txtMRP")) <> "" Then
		UNIValue(0,8) = "  " & FilterVar(Trim(UCase(Request("txtMRP"))), " " , "S") & "  "
	Else 
	    UNIValue(0,8) = "|"
	End If
	
	'tracking_no 200308
	If Trim(Request("txtTrackingNo")) <> "" Then
		UNIValue(0,9) = "  " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), " " , "S") & "  "
	Else 
	    UNIValue(0,9) = "|"
	End If
	
     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
	If SetConditionData = False Then Exit Sub

    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()

    End If  
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

'============================================================================================================
' Name : SubLookSppl
' Desc : 공급처 Change Event
'============================================================================================================
Sub SubLookSppl

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim iPM2G139
    Dim strPrNo
    Dim strSpplCd
    Dim iArrItemByPlant
    Dim iArrPurGrp
    Dim iArrSpplCal
    
    Const C_sppl_dvly_dt = 0
    Redim iArrItemByPlant(C_sppl_dvly_dt)

    Const C_pr_grp_cd = 0
    Const C_pr_grp_nm = 1
    Redim iArrPurGrp(C_pr_grp_nm)

    Const C_cal_dt = 0
    Redim iArrSpplCal(C_cal_dt)
    
    Set iPM2G139 = Server.CreateObject("PM2G139.cMLookupSpplLtS")
	
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
		If CheckSYSTEMError(Err, True) = True Then
			Set iPM2G139 = Nothing
			Exit Sub
		End If
	
	strSpplCd = Trim(Request("txtBpCd"))
	strPrNo = Trim(Request("txtPrNo"))

	Call iPM2G139.M_LOOKUP_SPPL_LT_SVR(gStrGlobalCollection, strPrNo, _
										strSpplCd, iArrItemByPlant, _
										iArrPurGrp, iArrSpplCal)
	
	
		If CheckSYSTEMError(Err, True) = True Then
			Set iPM2G139 = Nothing
			Exit Sub
		End If	
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " With Parent.frm1.vspdData2 "      & vbCr
    Response.Write " .Row  =  .ActiveRow "			  & vbCr
    Response.Write " .Col 	= Parent.C_GrpCd "        & vbCr
    Response.Write "   If .text = """" Then "         & vbCr	
    Response.Write "      .text   = """ & ConvSPChars(iArrPurGrp(C_pr_grp_cd)) & """" & vbCr	
    Response.Write "      .Col 	= Parent.C_GrpNm "    & vbCr	
    Response.Write "      .text   = """ & ConvSPChars(iArrPurGrp(C_pr_grp_nm)) & """" & vbCr	
    Response.Write "   End If "             & vbCr
    Response.Write " End With "             & vbCr	
    Response.Write "</Script> "            

	Set iPM2G139 = Nothing
End Sub
%>
