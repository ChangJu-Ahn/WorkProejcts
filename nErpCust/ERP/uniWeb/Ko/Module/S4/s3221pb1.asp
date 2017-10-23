<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3221pb1.asp																*
'*  4. Program Name         : L/C Amend번호(L/C Amend등록에서)	 										*
'*  5. Program Desc         : L/C Amend번호(L/C Amend등록에서)											*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/03																*
'*  8. Modified date(Last)  : 2002/07/09																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/03 : 화면 design												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%   

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")

On Error Resume Next

Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   '☜ : DBAgent Parameter 선언 
Dim rs1, rs2
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strApplicant	                                                       
Dim strSalesGroup
Dim BlankchkFlg	                                                           
Const C_SHEETMAXROWS_D  = 30      
Dim arrRsVal(3)			
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
Call HideStatusWnd 
	
lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)				   '☜:
lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag   
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
lgDataExist    = "No"
	
Call TrimData()
Call FixUNISQLData()
Call QueryData()
'----------------------------------------------------------------------------------------------------------
' Query Data
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

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                              '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    
	Dim strVal
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,2)

    UNISqlId(0) = "S3221PA101"									'* : 데이터 조회를 위한 SQL문 만듬 
	
	UNISqlId(1) = "S3211PA102"
	UNISqlId(2) = "S3211PA103"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
	UNIValue(1,0)  = UCase(Trim(strApplicant))
    UNIValue(2,0)  = UCase(Trim(strSalesGroup))
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = " "
    
	If Trim(Request("txtApplicant")) <> "" Then
		strVal = strVal& " AND D.APPLICANT >= " & FilterVar(Request("txtApplicant"), "''", "S") & "  AND D.APPLICANT <=  " & FilterVar(Request("txtApplicant"), "''", "S") & " "
	Else
		strVal = strVal& " AND D.APPLICANT >='' AND D.APPLICANT <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If

	If Trim(Request("txtLCDocNo")) <> "" Then
		strVal = strVal& " AND D.LC_DOC_NO >= " & FilterVar(Request("txtLCDocNo"), "''", "S") & "  AND D.LC_DOC_NO <=  " & FilterVar(Request("txtLCDocNo"), "''", "S") & " "
	Else
		strVal = strVal& " AND D.LC_DOC_NO >='' AND D.LC_DOC_NO <= " & FilterVar("zzzzzzzzzzzzzzzzzz", "''", "S") & " "
	End If


	If Trim(Request("txtSalesGroup")) <> "" Then
		strVal = strVal& " AND D.sales_grp >= " & FilterVar(Request("txtSalesGroup"), "''", "S") & "  AND D.sales_grp <=  " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "
	Else
		strVal = strVal& " AND D.sales_grp >='' AND D.sales_grp <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If
			
	
	If Len(Trim(Request("txtFromDt"))) Then
		strVal = strVal & " AND D.AMEND_DT >= " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToDt"))) Then
		strVal = strVal & " AND D.AMEND_DT <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""		
	End If

    UNIValue(0,1) = strVal   
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    
	Dim iStr
	Dim FalsechkFlg
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2) '* : Record Set 의 갯수 조정 
    
    
    FalsechkFlg = False
	
	'============================= 추가된 부분 =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtApplicant")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수입자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
			Response.Write "<Script language=vbs>  " & vbCr   
            Response.Write "   Parent.frm1.txtApplicant.focus " & vbCr    
            Response.Write "</Script>      " & vbCr	
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
			Response.Write "<Script language=vbs>  " & vbCr   
			Response.Write "   Parent.frm1.txtSalesGroup.focus " & vbCr    
			Response.Write "</Script>      " & vbCr	
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
    
	If Len(Request("txtLCDocNo")) = "" Then
		   Call DisplayMsgBox("970000", vbInformation, "L/C번호", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			BlankchkFlg = True	
	       	Response.Write "<Script language=vbs>  " & vbCr   
			Response.Write "   Parent.frm1.txtLCDocNo.focus " & vbCr    
			Response.Write "</Script>      " & vbCr	
	End If	

	
	

     iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    
    If BlankchkFlg = False Then
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
			Response.Write "<Script language=vbs>  " & vbCr   
            Response.Write "   Parent.frm1.txtApplicant.focus " & vbCr    
            Response.Write "</Script>      " & vbCr			    
		Else    
		    Call  MakeSpreadSheetData()
		End If
    End If


End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
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
  				.frm1.txtHLCDocNo.value		 =  "<%=ConvSPChars(Request("txtLCDocNo"))%>" 	
				.frm1.txtHFromDt.value		 =  "<%=ConvSPChars(Request("txtFromDt"))%>" 	
  				.frm1.txtHToDt.value		 =  "<%=ConvSPChars(Request("txtToDt"))%>" 	
  					
			End If
			.ggoSpread.Source    = .frm1.vspdData
			.frm1.vspdData.Redraw = False 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"           '☜: Display data
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,-1,-1,.GetKeyPos("A",2),.GetKeyPos("A",3),"A","Q","X","X") 
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag		
			.DbQueryOk
			.frm1.vspdData.Redraw = True
		End with
	
	End If   
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>

