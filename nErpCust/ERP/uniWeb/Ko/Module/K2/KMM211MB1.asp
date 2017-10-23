<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MM211MB1
'*  4. Program Name         : 멀티컴퍼니B/L등록 
'*  5. Program Desc         : 멀티컴퍼니B/L등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/07/02
'*  8. Modified date(Last)  : 2005/02/04
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : Moon Jeong GIL
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

'Response.Write gStrGlobalCollection
'Response.end
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
Dim SoCompanyNm		'☜ : 수주법인 
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
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear
	Dim iErrorPosition	
	Dim LngMaxRow
	Dim arrTemp
	Dim arrVal
	Dim lGrpCnt
	Dim LngRow
	Dim iRow_cnt 

	Dim iCFM_YN					'확정여부 
	Dim iPO_COMPANY				'발주법인 
	Dim iSO_COMPANY				'수주법인 
	Dim iBL_NO					'
	Dim itxtIV_TYPE_CD			'매입형태 
	Dim itxtLoadingDt			'선적일 
	Dim itxtBlIssueDt			'B/L접수일 
	
'	Dim itxtIV_DT				
	Dim itxtPUR_GRP				'구매그룹 
	Dim iCUST_PO_NO				'발주번호 
	
	Dim ObjPMMG211
	     

	LngMaxRow = CInt(Request("txtMaxRows"))								'☜: 최대 업데이트된 갯수 
	arrTemp = Split(Request("txtSpread"), gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 


	
	lGrpCnt = 0
	
	Set ObjPMMG211 = Server.CreateObject ("PMMG211.MMaintMcSpplBlSvr")    
	
	If CheckSYSTEMError(Err,True) = true then
		Set ObjPMMG211 = Nothing		
		Exit Sub
	End If	
	
	'//Response.Write "arrTemp(0):" & arrTemp(0) & "<br>"
	'//Response.Write "arrTemp(1):" & arrTemp(1) & "<br>"

	For LngRow = 1 To LngMaxRow
			Err.Clear
			
	
			arrVal = Split(arrTemp(LngRow-1), gColSep)
			
				
			iCFM_YN		= arrVal(7)
			iPO_COMPANY	= arrVal(18)														
			iSO_COMPANY	= arrVal(14)	
			iBL_NO 		= arrVal(8)	
			iCUST_PO_NO = arrVal(15)
	
			'매입형태 	txtIV_TYPE_CD
			itxtIV_TYPE_CD	= arrVal(2)	
			'선적일		txtLoadingDt
			itxtLoadingDt 	= arrVal(3)	
			'B/L접수일		txtBlIssueDt
			itxtBlIssueDt 	= arrVal(4)				
			'구매그룹	txtPUR_GRP
			itxtPUR_GRP	= arrVal(5)	
			
	
			'Response.write "--------------------------" &"<br>"
			'Response.write "iCFM_YN:" & iCFM_YN &"<br>"
			'Response.write "iPO_COMPANY:" & iPO_COMPANY &"<br>"
			'Response.write "iSO_COMPANY:" & iSO_COMPANY &"<br>"
			'Response.write "iBL_NO:" & iBL_NO &"<br>"
			'Response.write "itxtSo_Type:" & itxtSo_Type &"<br>"
			'Response.write "itxtDeal_Type:" & itxtDeal_Type &"<br>"
			'Response.write "itxtSales_Grp:" & itxtSales_Grp &"<br>"
			'Response.write "itxtPlantCd:" & itxtPlantCd &"<br>"
			
			'Response.write "--------------------------" &"<br>"
	
			On Error Resume Next                                                             '☜: Protect system from crashing
			Err.Clear
	
			Call ObjPMMG211.M_MAINT_MC_SPPL_BL_SVR(gStrGlobalCollection,	iCFM_YN, _
											iPO_COMPANY, _
											iSO_COMPANY, _
											iBL_NO, _
											itxtIV_TYPE_CD, _
											itxtLoadingDt, _
											itxtBlIssueDt, _
											itxtPUR_GRP, _
											iCUST_PO_NO, _
											iErrorPosition)
									
			'-----------------------
			'Com action result check area(DB,internal)
			'-----------------------
			If CheckSYSTEMError2(Err, True, LngRow & "행:", "", "", "", "") = True Then
			    	Err.Clear
				'처리가 완료된것은 Check Box 가 풀림.
				Response.Write "<Script language=vbscript> "		& vbCr  
				Response.Write "	Dim iBln "				& vbCr      
				Response.Write "            iBln = MsgBox (""계속진행하시겠습니까?"", vbYesNo, """") "				& vbCr      
				Response.Write "            If iBln = vbNo Then   "				& vbCr      
				Response.Write "	       Parent.DbSaveOk    "				& vbCr      
				Response.Write "	    End If"						& vbCr      
				Response.Write "</Script> "		    
			Else
				'처리가 완료된것은 Check Box 가 풀림.
				Response.Write "<Script language=vbscript> "		& vbCr  
				Response.Write "On error resume Next"				& vbCr      
				Response.Write "	with Parent.frm1.vspdData"      & vbCr	 			
				Response.Write "		Dim iIndex, iRowNo	"		& vbCr	
				Response.Write "		for iIndex = 1 to .MaxRows	"      & vbCr	
				Response.Write "			.Col = Parent.C_BL_NO	"      & vbCr
				Response.Write "			.Row = iIndex	"		& vbCr		
				Response.Write "			If Trim(.text) = """	&  iBL_NO & """ then "     & vbCr			
				Response.Write "				iRowNo = iIndex	"   & vbCr
				Response.Write "			End if	"				& vbCr	
				Response.Write "		Next	"					& vbCr	
				Response.Write "		.Col = parent.C_CfmFlg	"   & vbCr		
				Response.Write "		.Row = iRowNo "				& vbCr				
'				Response.Write "		.Text = 0 "					& vbCr		
				Response.Write "	end with "						& vbCr	
				Response.Write "</Script> "		    
			    
			End If			
		
	Next  
	
	If NOT(ObjPMMG211 is Nothing) Then
		Set ObjPMMG211 = Nothing	
	End If		

        
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "     

	       
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



'	Const	M_MC_SPPL_INV_LIST_H_CHECK				=	0		'선택                  
'	Const	M_MC_SPPL_INV_LIST_H_CFM_FLG		    =	1       '확정여부              
	Const	M_MC_SPPL_INV_LIST_H_BL_NO			    =	0       '공급처B/L번호         
	Const	M_MC_SPPL_INV_LIST_H_BL_DOC_AMT		    =	1       'B/L금액               
	Const	M_MC_SPPL_INV_LIST_H_BL_CUR			    =	2       '화폐                  
	Const	M_MC_SPPL_INV_LIST_H_BL_LOC_AMT		    =	3       'B/L자국금액           
	Const	M_MC_SPPL_INV_LIST_H_XCH_RT			    =	4       '환율                  
	Const	M_MC_SPPL_INV_LIST_H_BL_PAY_METH		=	5       '결제방법              
	Const	M_MC_SPPL_INV_LIST_H_BL_PAY_METH_NM	    =	6       '결재방법명			   
	Const	M_MC_SPPL_INV_LIST_H_BL_PAY_TYPE		=	7       '지급유형              
	                                                                                   
	Const	M_MC_SPPL_INV_LIST_H_PAYEE_CD		    =	8      '지급처                
	Const	M_MC_SPPL_INV_LIST_H_BUILD_CD		    =	9      '계산서발행처          
	Const	M_MC_SPPL_INV_LIST_H_BENEFICIARY	    =	10      '수출자                
	Const	M_MC_SPPL_INV_LIST_H_CUST_PO_NO		    =	11      '발주번호              
'	Const	M_MC_SPPL_INV_LIST_H_TRANSPORT		    =	12      '운송방법              
'	Const	M_MC_SPPL_INV_LIST_H_PAY_DUR		    =	13      '결제기간            
	Const	M_MC_SPPL_INV_LIST_H_C_PO_COMPANY	    =	12      '발주법인  
	Const	M_MC_SPPL_INV_LIST_H_BL_DOC_NO		    =	13      '발주법인  
  
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       intTRows		= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
	
	'----- 레코드셋 칼럼 순서 ----------
	'A.BL_NO, A.BL_DOC_AMT, A.BL_CUR, A.BL_LOC_AMT, A.XCH_RT, A.BL_PAY_METH, C.MINOR_NM BL_PAY_METH_NM,
	'A.BL_PAY_TYPE, A.SO_COMPANY PAYEE_CD, A.SO_COMPANY BUILD_CD, A.SO_COMPANY BENEFICIARY,
	'B.CUST_PO_NO, PO_COMPANY
	'-----------------------------------
	
	iLoopCount = 0
    ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		
		iRowStr = ""
  
		iRowStr = iRowStr & Chr(11) & "0"																					'선택                 
		iRowStr = iRowStr & Chr(11) & "0"		                                                                            '확정여부             
				        	
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_BL_NO))			                                '공급처B/L번호        
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_SPPL_INV_LIST_H_BL_DOC_AMT), ggAmtOfMoney.DecPoint,0)     'B/L금액              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_BL_CUR))			                                '화폐                 
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_SPPL_INV_LIST_H_BL_LOC_AMT), ggAmtOfMoney.DecPoint,0)	    'B/L자국금액          
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_SPPL_INV_LIST_H_XCH_RT), ggExchRate.DecPoint,6)			'환율                 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_BL_PAY_METH))	                                '결제방법             
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_BL_PAY_METH_NM))	                                '결재방법명			  
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_BL_PAY_TYPE))	                                '지급유형             
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_PAYEE_CD))		                                '지급처               
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_BUILD_CD))		                                '계산서발행처         
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_BENEFICIARY))	                                '수출자               
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_CUST_PO_NO))		                                '발주번호             
		iRowStr = iRowStr & Chr(11) & ""													                                '운송방법             
		iRowStr = iRowStr & Chr(11) & ""		                                											'결제기간                  
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_C_PO_COMPANY))	                                '발주법인 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_SPPL_INV_LIST_H_BL_DOC_NO))	                                '발주법인                          
		
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
    Redim UNIValue(1,4)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 

	strVal = ""
    UNISqlId(0) = "MM211MA101"
    UNISqlId(1) = "MM111MA103"		'수주법인조회 
    
	UNIValue(1,0) = "'zzzzzzzzzz'"            
    
    '수주법인조회 
    If Trim(Request("txtSoCompanyCd")) <> "" Then
		UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	    UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,0) = "|"
	End If
	
    '공급처B/L발행일 
    If Trim(Request("txtFrBillDt")) <> "" Then
		UNIValue(0,1) =  " '" & Trim(UniConvDate(Request("txtFrBillDt"))) & "' "	
    Else
        UNIValue(0,1) = "|"
	End If
			
    If Trim(Request("txtToBillDt")) <> "" Then
		UNIValue(0,2) =  " '" & Trim(UniConvDate(Request("txtToBillDt"))) & "' "	
    Else
        UNIValue(0,2) = "|"
	End If

    '공급처B/L번호 
    If Trim(Request("txtBlNo")) <> "" Then
		UNIValue(0,3) = " '"& FilterVar(Trim(UCase(Request("txtBlNo"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,3) = "|"
	End If	
		
    '발주번호 
    If Trim(Request("txtCustPoNo")) <> "" Then
		UNIValue(0,4) = " '"& FilterVar(Trim(UCase(Request("txtCustPoNo"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,4) = "|"
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
