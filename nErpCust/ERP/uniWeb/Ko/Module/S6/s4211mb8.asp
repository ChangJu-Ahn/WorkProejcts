<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S4211MA8
'*  4. Program Name         : 통관현황조회 
'*  5. Program Desc         : 통관현황조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/12
'*  9. Modifier (First)     : Cho Sung-Hyun
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/29 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO변환 
'**************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    
On Error Resume Next


Err.Clear


Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  
Dim lgstrRetMsg                                                            
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   
Dim rs1, rs2 ,rs3
Dim lgstrData                                                              
Dim lgStrPrevKey                                                      
Dim lgTailList                                                             
Dim lgSelectList
Dim lgSelectListDT
Dim strApplicant	                                                       
Dim strSalesGroup	                                                           
Dim strEdType	
Dim BlankchkFlg                                                           
Dim arrRsVal(5)								
Const C_SHEETMAXROWS_D  = 100               
                    
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
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
    Set rs0 = Nothing	                                              '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(4,2)

    UNISqlId(0) = "S4211MA801"									'* : 데이터 조회를 위한 SQL문 만듬 
	
	UNISqlId(1) = "S4211MA802"			'수입자 
	UNISqlId(2) = "S3211PA103"			'영업그룹 
	UNISqlId(3) = "S4211MA803"			'신고구분 

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
	UNIValue(1,0)  = UCase(Trim(strApplicant))
    UNIValue(2,0)  = UCase(Trim(strSalesGroup))
    UNIValue(3,0)  = UCase(Trim(strEdType))
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = " "
    
	If Trim(Request("txtApplicantCd")) <> "" Then
		strVal = strVal& " AND A.APPLICANT >= " & FilterVar(UCase(Request("txtApplicantCd")), "''", "S") & "  AND A.APPLICANT <=  " & FilterVar(UCase(Request("txtApplicantCd")), "''", "S") & " "
	Else
		strVal = strVal& " AND A.APPLICANT >='' AND A.APPLICANT <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If

	If Trim(Request("txtSalesGrpCd")) <> "" Then
		strVal = strVal& " AND A.SALES_GRP >= " & FilterVar(UCase(Request("txtSalesGrpCd")), "''", "S") & "  AND A.SALES_GRP <=  " & FilterVar(UCase(Request("txtSalesGrpCd")), "''", "S") & " "
	Else
		strVal = strVal& " AND A.SALES_GRP >='' AND A.SALES_GRP <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If

  	If Trim(Request("txtEdType")) <> "" Then
		strVal = strVal& " AND A.ED_TYPE >= " & FilterVar(UCase(Request("txtEdType")), "''", "S") & "  AND A.ED_TYPE <=  " & FilterVar(UCase(Request("txtEdType")), "''", "S") & " "
	Else
		strVal = strVal& " AND A.ED_TYPE >='' AND A.ED_TYPE <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If
	
		
	If Len(Trim(Request("txtFromDate"))) Then
		strVal = strVal & " AND A.IV_DT >= " & FilterVar(UNIConvDate(Request("txtFromDate")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToDate"))) Then
		strVal = strVal & " AND A.IV_DT <= " & FilterVar(UNIConvDate(Request("txtToDate")), "''", "S") & ""		
	End If

    UNIValue(0,1) = strVal   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag,rs0,rs1,rs2,rs3) '* : Record Set 의 갯수 조정 
    
    Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)
	
	'============================= 추가된 부분 =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtApplicantCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수입자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtApplicantCd.focus    
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
        If Len(Request("txtSalesGrpCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGrpCd.focus    
                </Script>
            <%	       
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

	If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtEdType")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "신고구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtEdType.focus    
                </Script>
            <%	       
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
 
    If BlankchkFlg = False Then
		If rs0.EOF And rs0.BOF Then
		   Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		   rs0.Close
		   Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtApplicantCd.focus    
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

	'---수입자 
    If Len(Trim(Request("txtApplicantCd"))) Then
    	strApplicant = " " & FilterVar(Request("txtApplicantCd"), "''", "S") & " "
    	
    Else
    	strApplicant = "''"
    End If
    '---그룹 
    If Len(Trim(Request("txtSalesGrpCd"))) Then
    	strSalesGroup = " " & FilterVar(Request("txtSalesGrpCd"), "''", "S") & " "
    Else
    	strSalesGroup = "''"
    End If
	
	'---신고구분 
    If Len(Trim(Request("txtEdType"))) Then
    	strEdType = " " & FilterVar(Request("txtEdType"), "''", "S") & " "
    Else
    	strEdType = "''"
    End If

End Sub


%>
<Script Language=vbscript>
    parent.frm1.txtApplicantNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  	parent.frm1.txtSalesGrpNm.value			=  "<%=ConvSPChars(arrRsVal(3))%>" 	
	parent.frm1.txtEdTypeNm.value			=  "<%=ConvSPChars(arrRsVal(5))%>" 		
	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.HApplicantCd.value	= "<%=ConvSPChars(Request("txtApplicantCd"))%>"
				.frm1.HSalesGrpCd.value		= "<%=ConvSPChars(Request("txtSalesGrpCd"))%>"
				.frm1.HEdType.value			= "<%=ConvSPChars(Request("txtEdType"))%>"
				.frm1.HEpType.value			= "<%=ConvSPChars(Request("txtEpType"))%>"
				.frm1.HExportType.value		= "<%=ConvSPChars(Request("txtExportType"))%>"
				.frm1.HFromDate.value		= "<%=Request("txtFromDate")%>"
				.frm1.HToDate.value			= "<%=Request("txtToDate")%>"
			End If
			.ggoSpread.Source    = .frm1.vspdData 
			
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,-1,-1,.GetKeyPos("A",3),.GetKeyPos("A",4),"A","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,-1,-1,.GetKeyPos("A",3),.GetKeyPos("A",5),"A","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,-1,-1,.parent.gCurrency,.GetKeyPos("A",6),"A","Q","X","X")
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag		
			.DbQueryOk
			 .frm1.vspdData.Redraw = True
		End with
	
	End If   
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
