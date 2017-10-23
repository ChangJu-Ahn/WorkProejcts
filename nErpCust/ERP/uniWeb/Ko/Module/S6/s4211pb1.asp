<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S4211PB1
'*  4. Program Name         : 통관관리번호 팝업 
'*  5. Program Desc         : 통관관리번호 팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : kim hyung suk
'* 10. Modifier (Last)      : Seo jin kyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%   

Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  
Dim lgstrRetMsg                                                            
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   
Dim rs1, rs2
Dim lgstrData                                                              
Dim lgStrPrevKey                                                    
Dim lgTailList                                                             
Dim lgSelectList
Dim lgSelectListDT
Dim strApplicant	                                                       
Dim strSalesGroup	                                                           
Dim BlankchkFlg
Dim arrRsVal(3)								
Const C_SHEETMAXROWS_D  = 30      

    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")
	
    Call HideStatusWnd 
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)					
    lgStrPrevKey   = Request("lgStrPrevKey")                                   
    lgSelectList   = Request("lgSelectList")                               
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)                 
    lgTailList     = Request("lgTailList")                                 
	lgDataExist    = "No"
	
	Call TrimData()
	Call FixUNISQLData()
	Call QueryData()
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  
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

    If iLoopCount < C_SHEETMAXROWS_D Then                                      
       lgPageNo = ""
    End If
    rs0.Close                                                       
    Set rs0 = Nothing	                                              
End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    
	Dim strVal
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,2)

    UNISqlId(0) = "S4211PA101"									'* : 데이터 조회를 위한 SQL문 만듬 
	
	UNISqlId(1) = "S3211PA102"
	UNISqlId(2) = "S3211PA103"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
	UNIValue(1,0)  = UCase(Trim(strApplicant))
    UNIValue(2,0)  = UCase(Trim(strSalesGroup))
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = " "
    
	If Trim(Request("txtApplicant")) <> "" Then
		strVal = strVal& " AND B_BIZ_PARTNER01.BP_CD >= " & FilterVar(Request("txtApplicant"), "''", "S") & "  AND B_BIZ_PARTNER01.BP_CD <=  " & FilterVar(Request("txtApplicant"), "''", "S") & " "
	Else
		strVal = strVal& " AND B_BIZ_PARTNER01.BP_CD >='' AND B_BIZ_PARTNER01.BP_CD <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If

	If Trim(Request("txtSalesGroup")) <> "" Then
		strVal = strVal& " AND B_SALES_GRP02.SALES_GRP >= " & FilterVar(Request("txtSalesGroup"), "''", "S") & "  AND B_SALES_GRP02.sales_grp <=  " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "
	Else
		strVal = strVal& " AND B_SALES_GRP02.sales_grp >='' AND B_SALES_GRP02.sales_grp <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If

  	If Trim(Request("txtIVNo")) <> "" Then
  		strVal = strVal& " AND S_CC_HDR03.IV_NO >= " & FilterVar(Request("txtIVNo"), "''", "S") & "  AND S_CC_HDR03.IV_NO <=  " & FilterVar(Request("txtIVNo"), "''", "S") & " "		
	Else
		strVal = strVal& " AND S_CC_HDR03.IV_NO >='' AND S_CC_HDR03.IV_NO <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If
			
	
	If Len(Trim(Request("txtFromDt"))) Then
		strVal = strVal & " AND S_CC_HDR03.IV_DT >= " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToDt"))) Then
		strVal = strVal & " AND S_CC_HDR03.IV_DT <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""		
	End If

    UNIValue(0,1) = strVal   

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    
	Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2) '* : Record Set 의 갯수 조정 
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
	Dim FalsechkFlg
    
    FalsechkFlg = False
	
	'============================= 추가된 부분 =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtApplicant")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수입자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
			%>
			<Script Language=vbscript>
			    parent.frm1.txtApplicant.focus
			</Script>	
			<% 
		End If
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtSalesGroup")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
			%>
			<Script Language=vbscript>
			    parent.frm1.txtSalesGroup.focus
			</Script>	
			<% 
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
		
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
			%>
			<Script Language=vbscript>
			    parent.frm1.txtApplicant.focus
			</Script>	
			<% 
		    Exit Sub
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	
End Sub
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

	'---사업장 
    If Len(Trim(Request("txtApplicant"))) Then
    	strApplicant = " " & FilterVar(Request("txtApplicant"), "''", "S") & " "    	
    Else
    	strApplicant = "''"
    End If
    '---품목 
    If Len(Trim(Request("txtSalesGroup"))) Then
    	strSalesGroup = " " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "
    Else
    	strSalesGroup = "''"
    End If


End Sub

%>
<Script Language=vbscript>
    parent.frm1.txtApplicantNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>" 	
	parent.frm1.txtSalesGroupNm.value			=  "<%=ConvSPChars(arrRsVal(3))%>" 	
	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHApplicant.value	 =  "<%=ConvSPChars(Request("txtApplicant"))%>" 	
  				.frm1.txtHSalesGroup.value   =  "<%=ConvSPChars(Request("txtSalesGroup"))%>" 	
  				.frm1.txtIVNo.value			 =  "<%=ConvSPChars(Request("txtIVNo"))%>" 	
				.frm1.txtHFromDt.value		 =  "<%=Request("txtFromDt")%>"
  				.frm1.txtHToDt.value		 =  "<%=Request("txtToDt")%>"
  				
			End If
			.frm1.vspdData.Redraw = False
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"          '☜: Display data 
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag		
			.frm1.vspdData.Redraw = True
			.DbQueryOk
		
		End with
	
	End If   
</Script>	
<%
    Response.End													
%>


