<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M5141RB1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Purchase Order Detail 참조 PopUp ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2002/04/23																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Kim Jae Soon																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/08 : Coding Start												*
'********************************************************************************************************
%>
	
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0       		   '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgStrData_1
Dim lgMaxCount                                                '☜ : Spread sheet 의 visible row 수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim lgPageNo_1

Dim strPtnBpNm												  ' 남품처명 
Dim strDNTypeNm												  ' 출하형태명 
Dim strSOTypeNm											      ' 수주타입명 
Dim gridNum													  ' 그리드 순서 확인 
Const C_VatYn		=	12
Const C_VatYnDsc	=	13

    Call HideStatusWnd 
    

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	'이성용 추가 
	lgPageNo_1       = UNICInt(Trim(Request("lgPageNo_1")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(Request("lgMaxCount"))             '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"

	Dim iPrevEndRow_B, iEndRow_B

	iPrevEndRow_B = 0
	iEndRow_B  = 0
	gridNum			= Request("txtGridNum")

	Call FixUNISQLData(gridNum)									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
 
 '----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount 
    '이성룡 추가 
    Dim iLoopCount_1                                                                    
    Dim iRowStr,iRowStr_1
    Dim ColCnt
    
    lgDataExist    = "Yes"

	
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    If gridNum = "A" then
    
	    iLoopCount_1 = -1
        
		lgstrData_1	   = ""
		
		Do while Not (rs0.EOF Or rs0.BOF)
   
		     iLoopCount_1 =  iLoopCount_1 + 1
		     iRowStr_1 = ""
		     
				For ColCnt = 0 To UBound(lgSelectListDT) - 1 
					iRowStr_1 = iRowStr_1 & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
				Next
			 
			 '이성용 
			 If iLoopCount_1 < lgMaxCount Then
			 'call svrmsgbox(iLoopCount_1  , vbinformation, i_mkscript)
		     'If iLoopCount < lgMaxCount Then
		        lgstrData_1 = lgstrData_1 & iRowStr_1 & Chr(11) & Chr(12)
		     Else
		        lgPageNo_1 = lgPageNo_1 + 1
		        Exit Do
		     End If
		     
		     rs0.MoveNext
		Loop
	Else
	
	Dim RsStrtmp
		
		iLoopCount = -1

		lgstrData      = ""

		If CDbl(lgPageNo) > 0 Then
			iPrevEndRow_B = CDbl(lgMaxCount) * CDbl(lgPageNo_B)    
		End If 
		Do while Not (rs0.EOF Or rs0.BOF)
   
		     iLoopCount =  iLoopCount + 1
		     iRowStr = ""
		     
				For ColCnt = 0 To UBound(lgSelectListDT) - 1 
				
					If ColCnt = C_VatYnDsc-1 then
						If FormatRsString(lgSelectListDT(ColCnt),rs0(C_VatYn-1)) = "2" then
							RsStrtmp = "포함"
						else
							RsStrtmp = "별도"
						end if
					else
							RsStrtmp = FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
					end if
					
		         iRowStr = iRowStr & Chr(11) & RsStrtmp
				Next
 
		     If iLoopCount < lgMaxCount Then
		        lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
			     iEndRow_B = iPrevEndRow_B + iLoopCount + 1
		     Else
		        lgPageNo = lgPageNo + 1
			     iEndRow_B = iPrevEndRow_B + iLoopCount
		        Exit Do
		     End If
		     
		     rs0.MoveNext
		Loop
	End if
	
    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    If iLoopCount_1 < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo_1 = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub
   
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData(byVal gridNum)
	
    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(0,2)
    
	strVal = " "

	If Len(Request("txtFrIvDt")) Then
		strVal = strVal & " AND A.IV_DT >= '" & UNIConvDate(Request("txtFrIvDt")) & "' "	
	else
	    strVal = strVal & " AND A.IV_DT >= '1900-01-01' "			
	End If

	If Len(Request("txtToIvDt")) Then
		strVal = strVal & " AND A.IV_DT <= '" & UNIConvDate(Request("txtToIvDt")) & "' "	
	else
	    strVal = strVal & " AND A.IV_DT <= '2999-12-31' "			
	End If	

 	If Len(Request("txtIvTypeCd")) Then
			strVal = strVal & " AND A.IV_TYPE_CD = " & FilterVar(Trim(UCase(Request("txtIvTypeCd"))), " " , "S") & " "		
	End If	    
 	
 	If Len(Request("txtGrpCd")) Then
			strVal = strVal & " AND A.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtGrpCd"))), " " , "S") & " "		
	End If	    
	
    If Len(Request("txtVatCd")) Then
		strVal = strVal & " AND A.VAT_TYPE = " & FilterVar(Trim(UCase(Request("txtVatCd"))), " " , "S") & " "		
	End If	
	    
	If Len(Request("txtSpplCd")) Then
		strVal = strVal & " AND A.BP_CD = " & FilterVar(Trim(UCase(Request("txtSpplCd"))), " " , "S") & " "		
	End If
	
	If Len(Request("txtBuildCd")) Then
		strVal = strVal & " AND A.BUILD_CD = " & FilterVar(Trim(UCase(Request("txtBuildCd"))), " " , "S") & " "		
	End If		
	
	If Len(Request("txtPoNo")) Then
		strVal = strVal & " AND J.PO_NO like " & FilterVar(Trim(UCase(Request("txtPoNo")&"%")), " " , "S") & " "		
	End If		

	If Len(Request("txtIvNo")) Then
		strVal = strVal & " AND A.IV_NO like " & FilterVar(Trim(UCase(Request("txtIvNo")&"%")), " " , "S") & " "		
	End If
	
	If Len(Request("txtCur")) Then
		strVal = strVal & " AND A.IV_CUR like " & FilterVar(Trim(UCase(Request("txtCur")&"%")), " " , "S") & " "		
	End if
	
	'미확정건만 참조하도록 조건부 추가 
		strVal = strVal & " AND A.POSTED_FLG = 'Y' "
	
	if gridNum = "A" then
		UNISqlId(0) = "M5141RA1_1"
		
		UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
		UNIValue(0,1) = strVal & "GROUP BY A.IV_TYPE_CD , B.IV_TYPE_NM , A.BUILD_CD , C.BP_FULL_NM , A.PAYEE_CD , D.BP_FULL_NM , A.BP_CD , E.BP_FULL_NM , A.PUR_GRP , F.PUR_GRP_NM , A.TAX_BIZ_AREA , G.TAX_BIZ_AREA_NM , A.IV_CUR , A.VAT_TYPE , H.MINOR_NM , A.VAT_RT , A.PAY_METH , I.MINOR_NM , A.IV_DT , C.BP_RGST_NO , A.SPPL_IV_NO , A.PAY_DUR , A.PAY_TYPE , K.MINOR_NM , A.PAY_TERMS_TXT , A.REMARK  " & UCase(Trim(lgTailList)) 
					
    else
    
		If len(Request("txtPayeeCd")) Then
			strVal = strVal & " AND A.PAYEE_CD = " & FilterVar(Trim(UCase(Request("txtPayeeCd"))), " " , "S") & " "		
		End If
		If len(Request("txtCurr")) Then
			strVal = strVal & " AND A.IV_CUR = " & FilterVar(Trim(UCase(Request("txtCurr"))), " " , "S") & " "		
		End If
		If len(Request("txtPayTermCd")) Then
			strVal = strVal & " AND A.PAY_METH = " & FilterVar(Trim(UCase(Request("txtPayTermCd"))), " " , "S") & " "		
		End If
	
		UNISqlId(0) = "M5141RA1_2"
	
		UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
		UNIValue(0,1) = strVal & UCase(Trim(lgTailList)) 
		
	End if
	
	
    
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

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
    
        
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
       
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
    
    Set lgADF   = Nothing
    
End Sub



%>

<Script Language=vbscript>
    With parent
		
		If "<%= gridNum %>" = "A" Then
			If "<%=lgDataExist%>" = "Yes" Then
				.frm1.hdnFrDt.Value 		= "<%=Request("txtFrIvDt")%>"
				.frm1.hdnToDt.Value 		= "<%=Request("txtToIvDt")%>"
				.frm1.hdnIvTypeCd.Value 	= "<%=ConvSPChars(Request("txtIvTypeCd"))%>"
				.frm1.hdnGrpCd.Value 		= "<%=ConvSPChars(Request("txtGrpCd"))%>"
				.frm1.hdnVatCd.value		= "<%=ConvSPChars(Request("txtVatCd"))%>"			
				.frm1.hdnSpplCd.value		= "<%=ConvSPChars(Request("txtSpplCd"))%>"	
				.frm1.hdnBuildCd.value		= "<%=ConvSPChars(Request("txtBuildCd"))%>"
				.frm1.hdnPoNo.value			= "<%=ConvSPChars(Request("txtPoNo"))%>"
				.frm1.hdnIvNo.value			= "<%=ConvSPChars(Request("txtIvNo"))%>"
				.ggoSpread.Source			= .frm1.vspdData1 
				.ggoSpread.SSShowData "<%=lgstrData_1%>"                            '☜: Display data 


			' 이성용 수정 
				'.lgPageNo_1			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
				.lgPageNo_1			 =  "<%=lgPageNo_1%>"							  '☜: Next next data tag
				.DbQueryOk
			End If
		Else
			If "<%=lgDataExist%>" = "Yes" Then

				.ggoSpread.Source    = .frm1.vspdData2 
				.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 

				Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData2, <%=iPrevEndRow_B+1%>, <%=iEndRow_B%>, .frm1.hdnCurr1.value, .GetKeyPos("B",10),"C","Q","X","X") ' 매입단가 
				Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData2, <%=iPrevEndRow_B+1%>, <%=iEndRow_B%>, .frm1.hdnCurr1.value, .GetKeyPos("B",11),"A","Q","X","X") ' 매입금액 
				Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData2, <%=iPrevEndRow_B+1%>, <%=iEndRow_B%>, .frm1.hdnCurr1.value, .GetKeyPos("B",14),"A","Q","X","X") ' VAT금액 
				Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData2, <%=iPrevEndRow_B+1%>, <%=iEndRow_B%>,  "<%=gCurrency%>", .GetKeyPos("B",15),"A","Q","X","X") ' 매입자국금액 
				Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData2, <%=iPrevEndRow_B+1%>, <%=iEndRow_B%>,  "<%=gCurrency%>", .GetKeyPos("B",16),"A","Q","X","X") ' VAT자국금액 

				.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
				

				.DbQuery2Ok
			End If
		End  if
	End with
</Script>