<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MM212MB1
'*  4. Program Name         : 멀티컴퍼니B/L확정/삭제-멀티 
'*  5. Program Desc         : 멀티컴퍼니B/L확정/삭제-멀티 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/01/24
'*  8. Modified date(Last)  : 2005/03/07
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : MJG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :          :
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
Dim SoCompanyNm			'☜ : 수주법인 
Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows

Dim iStrPostingFlg

intARows=0
intTRows=0
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status


Call HideStatusWnd                                                               '☜: Hide Processing message
lgOpModeCRUD  = Request("txtMode") 

'Call ServerMesgBox(lgOpModeCRUD , vbInformation, I_MKSCRIPT)

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
		 Call  SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)
         Call SubBizDelete()
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

'	Call DisplayMsgBox(lgStrSQL, vbInformation, "", "", I_MKSCRIPT)


	Call FixUNISQLData()		'☜ : DB-Agent로 보낼 parameter 데이타 set
	
	Call QueryData()			'☜ : DB-Agent를 통한 ADO query
	
	'-----------------------
	'Result data display area
	'----------------------- 

%>

	<Script Language=vbscript>
		With parent
			.frm1.txtSoCompanyCd.value = "<%=ConvSPChars(Request("txtSoCompanyCd"))%>"			
			.frm1.txtSoCompanyNm.Value	= "<%=SoCompanyNm%>"							
			.frm1.txtSoCompanyCd.focus
			
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

	On Error Resume Next
    Err.Clear																				'☜: Protect system from crashing
	
	Dim iPMMG212,iErrorPosition
	Dim iMaxRow, istrVal
   
	'-------------------
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount, ii

    Dim iCUCount
    Dim iDCount
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")
    
'    Call ServerMesgBox(itxtSpread , vbInformation, I_MKSCRIPT)	
    
    Set iPMMG212 = Server.CreateObject("PMMG212.cMMaintBlCombi")        
	    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
			Set iPMMG212 = Nothing 		
			Exit Sub
	End If
	    
	
	Call iPMMG212.M_MAINT_BL_COMBI_SVR("F", gStrGlobalCollection, _
									  iCUCount, _
									  itxtSpread, _
									  iErrorPosition)								  	    									  

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iPMMG212 = Nothing												'☜: ComProxy Unload
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If
	 
	 

    Set iPMMG212 = Nothing															'☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Call parent.DbSaveOk()" & vbCr
	Response.Write "</Script>" & vbCr		
	

End Sub


'============================================================================================================
' Name : SubBizDelete
' Desc : Save Data into Db
'============================================================================================================

Sub SubBizDelete()																			'☜: 삭제 요청 
	
	On Error Resume Next
    Err.Clear																				'☜: Protect system from crashing
	
	Dim iPMMG212,iErrorPosition
	Dim iMaxRow, istrVal
   

	iMaxRow										= Trim(Request("txtMaxRows"))
	istrVal										= Trim(Request("txtSpread"))	
'	Call ServerMesgBox(istrVal , vbInformation, I_MKSCRIPT)			   
    
    Set iPMMG212 = Server.CreateObject("PMMG212.cMMaintBlCombi")        
	    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
			Set iPMMG212 = Nothing 		
			Exit Sub
	End If
	    
	
	Call iPMMG212.M_MAINT_BL_COMBI_SVR("F", gStrGlobalCollection, _
									  iMaxRow, _
									  istrVal, _
									  iErrorPosition)								  	    									  

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iPMMG212 = Nothing												'☜: ComProxy Unload
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If
	 
	 

    Set iPMMG212 = Nothing															'☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Call parent.DbSaveOk()" & vbCr
	Response.Write "</Script>" & vbCr															'☜: Unload Comproxy
												
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

		Const	M_IV_DTL_BUILD_CD			=	0
		Const	B_BIZ_PARTNE_BP_NM          =	1
		Const	M_IV_DTL_BL_NO              =	2
		Const	M_IV_DTL_BL_DOC_NO          =	3
		Const	M_IV_DTL_POSTING_FLG        =	4
		Const	M_IV_DTL_BL_ISSUE_DT        =	5
		Const	M_IV_DTL_LOADING_DT         =	6
		Const	M_IV_DTL_CURRENCY           =	7
		Const	M_IV_DTL_DOC_AMT            =	8
		Const	M_IV_DTL_LOC_AMT            =	9
		Const	M_IV_DTL_XCH_RATE           =	10
		Const	M_IV_DTL_IV_TYPE            =	11
		Const	M_IV_DTL_IV_TYPE_NM         =	12
		Const	M_IV_DTL_PAY_METHOD         =	13
		Const	M_IV_DTL_PAY_METHOD_NM      =	14
		Const	M_IV_DTL_PUR_GRP            =	15
		Const	M_IV_DTL_PUR_GRP_NM         =	16
		Const	M_IV_DTL_BENEFICIARY        =	17
		Const	M_IV_DTL_APPLICANT          =	18
		Const	M_IV_DTL_REF_IV_NO          =	19
		    
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       intTRows		= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
	
	'----- 레코드셋 칼럼 순서 ----------
	'A.BUILD_CD, (SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD =  A.BUILD_CD) BP_NM, A.BL_NO, A.BL_DOC_NO,
	'A.POSTING_FLG, A.BL_ISSUE_DT, A.LOADING_DT, A.CURRENCY, A.DOC_AMT, A.LOC_AMT, A.XCH_RATE, A.IV_TYPE,
	'(SELECT IV_TYPE_NM FROM M_IV_TYPE WHERE IV_TYPE_CD = A.IV_TYPE) IV_TYPE_NM, A.PAY_METHOD,
	'(SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'B9004' AND MINOR_CD = A.PAY_METHOD) PAY_METHOD_NM,
	'A.PUR_GRP, (SELECT PUR_GRP_NM FROM B_PUR_GRP WHERE PUR_GRP = A.PUR_GRP) PUR_GRP_NM
	'-----------------------------------

	iLoopCount = 0
    ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_BUILD_CD))												'수주법인              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_BIZ_PARTNE_BP_NM))											    '수주법인명            
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_BL_DOC_NO))											    '매입번호              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_BL_NO))										    '공급처세금계산서번호  
		IF rs0(M_IV_DTL_POSTING_FLG) = "Y" Then
			iStrPostingFlg = "확정"
		Else
			iStrPostingFlg = "미확정"			
		End if
		
		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrPostingFlg)	                                        '매입확정여부          
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_BL_ISSUE_DT))	                                            '매입일                
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_LOADING_DT))	                                        '매입그룹              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_CURRENCY))	                                        '매입그룹명            
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_DTL_DOC_AMT), ggAmtOfMoney.DecPoint,0)                                           '화폐                  
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_DTL_LOC_AMT), ggAmtOfMoney.DecPoint,0)	    '매입순금액            	
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_DTL_XCH_RATE), ggExchRate.DecPoint,6)			    '부가세율              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(Ucase(rs0(M_IV_DTL_IV_TYPE)))										    '부가세유형            
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_IV_TYPE_NM))									    '부가세유형명          
		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_PAY_METHOD))										'가격조건              	              		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_PAY_METHOD_NM))									'가격조건명            	    		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(Ucase(rs0(M_IV_DTL_PUR_GRP)))										    '결제방법              		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_PUR_GRP_NM))									    '결제방법명            	         				
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_BENEFICIARY))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_APPLICANT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_POSTING_FLG))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_DTL_REF_IV_NO))
		
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
	If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
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
		SoCompanyNm = rs1("BP_NM")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtSoCompanyCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "수주법인", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
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
	Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(1,5)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 

	strVal = ""
    UNISqlId(0) = "MM212MA101"
    UNISqlId(1) = "MM111MA103"		'수주법인조회 
    
	UNIValue(1,0) = "'zzzzzzzzzz'"            
    
    '수주법인조회 
    If Trim(Request("txtSoCompanyCd")) <> "" Then
		UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	    UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,0) = "|"
	End If
	
    'B/L접수일 
    If Trim(Request("txtBlIssueFrDt")) <> "" Then
		UNIValue(0,1) =  " '" & Trim(UniConvDate(Request("txtBlIssueFrDt"))) & "' "	
    Else
        UNIValue(0,1) = "|"
	End If
			
    If Trim(Request("txtBlIssueToDt")) <> "" Then
		UNIValue(0,2) =  " '" & Trim(UniConvDate(Request("txtBlIssueToDt"))) & "' "	
    Else
        UNIValue(0,2) = "|"
	End If	
	
    'B/L확정처리여부 
    If Trim(Request("rdoCfmflg")) <> "" Then
		UNIValue(0,3) = " '"& FilterVar(Trim(UCase(Request("rdoCfmflg"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,3) = "|"
	End If	
	
    '발주번호 
    If Trim(Request("txtPoNo")) <> "" Then
		UNIValue(0,4) = " '"& FilterVar(Trim(UCase(Request("txtPoNo"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,4) = "|"
	End If
	
    'B/L번호 
    If Trim(Request("txtBlNo")) <> "" Then
		UNIValue(0,5) = " '"& FilterVar(Trim(UCase(Request("txtBlNo"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,5) = "|"
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
        
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    	
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
