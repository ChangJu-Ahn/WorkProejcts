<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4111rb4.asp																*
'*  4. Program Name         : Local L/C 출하참조(Local L/C등록에서)										*
'*  5. Program Desc         : Local L/C 출하참조(Local L/C등록에서)										*
'*  6. Comproxy List        : S41118ListDnHdrForLcSvr													*
'*  7. Modified date(First) : 2000/10/11																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
																				
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
																		
On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				  '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3											  '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim strApplicantNm											  ' 주문처명 
Dim strSalesGroupNm											  ' 영업그룹명 
Dim strPayTermsNm
Dim BlankchkFlg										      ' 결제방법 

Dim iFrPoint
iFrPoint=0
Const C_SHEETMAXROWS_D  = 30                                          '☆: Fetch max count at once

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    	
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist    = "No"
	    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < C_SHEETMAXROWS_D Then                                 '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
    On Error Resume Next
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strApplicantNm =  rs1("BP_NM")
        rs1.Close
        Set rs1 = Nothing
    Else
		rs1.Close
		Set rs1 = Nothing
		If Len(Request("txtApplicant")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "개설신청인", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			BlankchkFlg = True
		%>
            <Script language=vbs>
            Parent.txtApplicant.focus    
            </Script>
         <%					
		End If
	End If   	
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strSalesGroupNm =  rs2("SALES_GRP_NM")
        rs2.Close
        Set rs2 = Nothing
    Else
		rs2.Close
		Set rs2 = Nothing
		If Len(Request("txtSalesGroup")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    BlankchkFlg = True
		%>
            <Script language=vbs>
            Parent.txtSalesGroup.focus    
            </Script>
         <%				    
		End If			
    End If   	
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strPayTermsNm =  rs3("MINOR_NM")
        rs3.Close
        Set rs3 = Nothing
    Else
		rs3.Close
		Set rs3 = Nothing
		If Len(Request("txtPayTerms")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "결제방법", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    BlankchkFlg = True
		%>
            <Script language=vbs>
            Parent.txtPayTerms.focus    
            </Script>
         <%				    
		End If				
    End If      
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,2)

    UNISqlId(0) = "S4111RA401"
    UNISqlId(1) = "s0000qa002"					'주문처명 
    UNISqlId(2) = "s0000qa005"					'영업그룹명 
    UNISqlId(3) = "s0000qa000"					'결제방법명  

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = ""

	If Len(Request("txtApplicant")) Then
		strVal = strVal & "AND c.bp_cd = " & FilterVar(Request("txtApplicant"), "''", "S") & "  "	
		arrVal(0) = Trim(Request("txtApplicant")) 
	End If

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & "AND d.sales_grp = " & FilterVar(Request("txtSalesGroup"), "''", "S") & "  "		
		arrVal(1) = Trim(Request("txtSalesGroup")) 
	End If		
		
 	If Len(Request("txtFromDt")) Then
		strVal = strVal & "AND b.so_dt >= " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & " "			
	End If		
	
	If Len(Request("txtToDt")) Then
		strVal = strVal & "AND b.so_dt <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & " "		
	End If
	
	If Len(Request("txtPayTerms")) Then
		strVal = strVal & "AND b.pay_meth = " & FilterVar(Request("txtPayTerms"), "''", "S") & "  "		
		arrVal(2) = Trim(Request("txtPayTerms")) 
	End If
	
	If Len(Request("txtRadio")) Then
		strVal = strVal & "AND b.lc_flag = " & FilterVar(Request("txtRadio"), "''", "S") & " "		
	End If				
    
	UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(Trim(Request("txtApplicant")), " " , "S")					'주문처코드 
    UNIValue(2,0) = FilterVar(Trim(Request("txtSalesGroup")), " " , "S")					'영업그룹코드 
    UNIValue(3,0) = FilterVar("B9004", " " , "S")						'결제방법코드    
    UNIValue(3,1) = FilterVar(Trim(Request("txtPayTerms")), " " , "S")					'결제방법코드    
    
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
  
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    BlankchkFlg = False
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

    Call  SetConditionData()


	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If BlankchkFlg = False Then         
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
		%>
            <Script language=vbs>
            Parent.txtApplicant.focus    
            </Script>
         <%		    
		Else    
		    Call  MakeSpreadSheetData()	    
		End If
	End If

End Sub

%>

<Script Language=vbscript>
	With parent
		.txtApplicantNm.value	= "<%=ConvSPChars(strApplicantNm)%>" 
		.txtSalesGroupNm.value	= "<%=ConvSPChars(strSalesGroupNm)%>" 
        .txtPayTermsNm.value	= "<%=ConvSPChars(strPayTermsNm)%>"
        
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.txtHApplicant.value	= "<%=ConvSPChars(Request("txtApplicant"))%>"
				.txtHSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>" 
				.txtHFromDt.value		= "<%=ConvSPChars(Request("txtFromDt"))%>"
				.txtHToDt.value			= "<%=ConvSPChars(Request("txtToDt"))%>"
				.txtHPayTerms.value		= "<%=ConvSPChars(Request("txtPayTerms"))%>"
			.DbQueryOk
			End If    
			'Show multi spreadsheet data from this line
			.ggoSpread.Source    = .vspdData 
			.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"
			
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,"<%=iFrPoint+1%>",parent.vspddata.maxrows,Parent.GetKeyPos("A",6),Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")
				
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			.DbQueryOk
			.vspdData.Redraw = True
		End If
	End with
</Script>
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
