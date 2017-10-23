<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        : 																			*
'*  3. Program ID           : s3211rb2.asp																*
'*  4. Program Name         : Local L/C참조(Local L/C Amend등록에서)									*
'*  5. Program Desc         : Local L/C참조(Local L/C Amend등록에서)									*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/04																*
'*  8. Modified date(Last)  : 2002/04/25																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Seo Jinkung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/04 : 화면 design												*
'*                            2. 2002/04/25 : Ado 변환													*
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
Call LoadBNumericFormatB("I","*","NOCOOKIE","RB")

On Error Resume Next	
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3, rs4		   '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iFrPoint
iFrPoint=0

Dim strApplicantNm			'개설신청인 
Dim strSalesGrpNm			'영업그룹 
Dim strCurNm				'화폐 
Dim strOpenBankNm			'개설은행 
Dim BlankchkFlg

Const C_SHEETMAXROWS_D  = 30                                          '☆: Fetch max count at once

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
 	
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
Sub SetConditionData()
    On Error Resume Next
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strApplicantNm =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtApplicant")) Then
			Call DisplayMsgBox("970000", vbInformation, "개설신청인", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
 %>
<Script Language=VBScript>
			Parent.frm1.txtApplicant.focus 
</Script>
<%		  		
		    BlankchkFlg = True
		    Response.End
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strSalesGrpNm =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtSalesGroup")) Then
			Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
 %>
<Script Language=VBScript>
			Parent.frm1.txtSalesGroup.focus 
</Script>
<%		  			
		    BlankchkFlg = True
		    Response.End
		End If			
    End If   	
    

    If Not(rs4.EOF Or rs4.BOF) Then        
        Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Request("txtCurrency")) Then
			Call DisplayMsgBox("970000", vbInformation, "화폐", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
 %>
<Script Language=VBScript>
			Parent.frm1.txtCurrency.focus 
</Script>
<%		  								
		    BlankchkFlg  =  True
		    Response.End
		End If				
    End If      

    
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strOpenBankNm =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtOpenBank")) Then
			Call DisplayMsgBox("970000", vbInformation, "개설은행", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
 %>
<Script Language=VBScript>
			Parent.frm1.txtOpenBank.focus 
</Script>
<%		  			
		    BlankchkFlg = True
		    Response.End
		End If				
    End If      
    

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(3)
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(5,2)

    UNISqlId(0) = "S3211RA201"
    UNISqlId(1) = "s0000qa002"					'개설신청인 
    UNISqlId(2) = "s0000qa005"					'영업그룹 
    UNISqlId(3) = "s0000qa008"					'개설은행  
    UNISqlId(4) = "s0000qa014"  ' 화폐     arrVal(3)
    
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
			
	strVal = " "	
	If Len(Request("txtApplicant")) Then
		strVal = "AND a.applicant = " & FilterVar(Request("txtApplicant"), "''", "S") & " "	
		arrVal(0) = Trim(Request("txtApplicant"))
	Else
		strVal = ""
	End If

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " AND a.sales_grp = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "		
		arrVal(1) = Trim(Request("txtSalesGroup"))
	End If		
		   
	If Len(Request("txtLCDocNo")) Then
		strVal = strVal & " AND a.lc_doc_no = " & FilterVar(Request("txtLCDocNo"), "''", "S") & " "				
	End If		
    
    If Len(Request("txtCurrency")) Then
		strVal = strVal & " AND a.cur  = " & FilterVar(Request("txtCurrency"), "''", "S") & " "				
		arrVal(3) = Trim(Request("txtCurrency"))
	End If		
	
	If Len(Request("txtOpenBank")) Then
		strVal = strVal & " AND a.issue_bank_cd  = " & FilterVar(Request("txtOpenBank"), "''", "S") & " "		
		arrVal(2) = Trim(Request("txtOpenBank"))
	End If			
 	   
	
    If Len(Request("txtFromDt")) Then
		strVal = strVal & " AND a.open_dt >= " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""			
	End If		
	
	If Len(Request("txtToDt")) Then
		strVal = strVal & " AND a.open_dt <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""		
	End If	
	

    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(Trim(Request("txtApplicant")), " " , "S")					'개설신청인 
    UNIValue(2,0) = FilterVar(Trim(Request("txtSalesGroup")), " " , "S")					'영업그룹    
    UNIValue(3,0) = FilterVar(Trim(Request("txtOpenBank")), " " , "S")					'개설은행 
    UNIValue(4,0) = FilterVar(Trim(Request("txtCurrency")), " " , "S") '	UNISqlId(4)의 첫번째 ? 
    
    
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

    Call  SetConditionData()

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    
    If BlankchkFlg = False Then         
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
 %>
<Script Language=VBScript>
			Parent.frm1.txtApplicant.focus 
</Script>
<%		    
		    rs0.Close
		    Set rs0 = Nothing
		Else    
		    Call  MakeSpreadSheetData()	    
		End If
	End If

    
    
End Sub

%>
<Script Language=VBScript>


	With parent
		.frm1.txtApplicantNm.value =  "<%=ConvSPChars(strApplicantNm)%>"
		.frm1.txtSalesGroupNm.value = "<%=ConvSPChars(strSalesGrpNm)%>"		
		.frm1.txtOpenBankNm.value = "<%=ConvSPChars(strOpenBankNm)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				parent.frm1.txtHApplicant.value = "<%=ConvSPChars(Request("txtApplicant"))%>"
				parent.frm1.txtHSalesGroup.value = "<%=ConvSPChars(Request("txtSalesGroup"))%>"
				parent.frm1.txtHLCDocNo.value = "<%=ConvSPChars(Request("txtLCDocNo"))%>"
				parent.frm1.txtHCurrency.value = "<%=ConvSPChars(Request("txtCurrency"))%>"
				parent.frm1.txtHOpenBank.value = "<%=ConvSPChars(Request("txtOpenBank"))%>"
				parent.frm1.txtHFromDt.value = "<%=Request("txtFromDt")%>"
				parent.frm1.txtHToDt.value = "<%=Request("txtToDt")%>"		
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"
			
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.GetKeyPos("A",6),Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")
				    	    
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = True
		End If
	End with   
	
</Script>
