<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : S3212RB6																	*
'*  4. Program Name         : L/C 내역참조(통관내역등록에서)											*
'*  5. Program Desc         : L/C 내역참조(통관내역등록에서)											*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/07																*
'*  8. Modified date(Last)  : 2002/04/																	*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kwak Eunkyoung															*
'* 11. Modifier             : SON BUM YEOL																*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/07 : 화면 design												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3           
   Dim lgStrData                                                  
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   Dim arrRsVal(1)
   Dim BlankchkFlg  
   
'--------------- 개발자 coding part(변수선언,Start)----------------------------------------------------
   Dim iFrPoint
   iFrPoint=0
   Const C_SHEETMAXROWS_D  = 30   
'--------------- 개발자 coding part(변수선언,End)------------------------------------------------------
	On Error Resume Next
	Err.Clear
	
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
    Call HideStatusWnd 
	
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)                  
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")	
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         

    lgDataExist      = "No"
	
    Call  FixUNISQLData()                                                
    call  QueryData()                                                    
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                
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

    If iLoopCount < C_SHEETMAXROWS_D Then                                      
       lgPageNo = ""
    End If
    rs0.Close                                                       
    Set rs0 = Nothing	                                            

End Sub
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()

End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(2)
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    'Redim UNIValue(1,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
    Redim UNIValue(1,3)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 

     UNISqlId(0) = "s3212ra601" 
     UNISqlId(1) = "s0000qa001" '품목팝업 

     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	strVal = " "

	If Len(Request("txtSONo")) Then
		strVal = strVal& " A.SO_NO =  " & FilterVar(Request("txtSONo"), "''", "S") & "  "
	Else
		strVal = strVal& " A.SO_NO >='' AND A.SO_NO <= " & FilterVar("zzzzzzzzz", "''", "S") & " "
	End If

	If Len(Request("txtItem")) Then
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(Trim(Request("txtItem")), "" , "S") & " "
		arrVal(0) = Trim(Request("txtItem"))
	End If

	If Len(Request("txtApplicant")) Then
		strVal = strVal & "AND A.APPLICANT = " & FilterVar(Request("txtApplicant"), "''", "S") & " "
	End If

	If Len(Request("txtCurrency")) Then
		strVal = strVal & " AND A.CUR = " & FilterVar(Request("txtCurrency"), "''", "S") & " "
	End If

    If Len(Trim(Request("txtSalesGroup"))) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "		
	End If		

    If Len(Trim(Request("txtPayTerms"))) Then
		strVal = strVal & " AND A.PAY_METH = " & FilterVar(Request("txtPayTerms"), "''", "S") & " "		
	End If		
	
    If Len(Trim(Request("txtIncoTerms"))) Then
		strVal = strVal & " AND A.INCOTERMS = " & FilterVar(Request("txtIncoTerms"), "''", "S") & " "		
	End If			
	
    If Len(Trim(Request("txtFromDt"))) Then
		strVal = strVal & " AND CONVERT(DATETIME, A.OPEN_DT) >= " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToDt"))) Then
		strVal = strVal & " AND CONVERT(DATETIME, A.OPEN_DT) <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""		
	End If
   
	If Len(Trim(Request("txtTrackingNo"))) Then
		strVal = strVal & " AND B.TRACKING_NO = " & FilterVar(Request("txtTrackingNo"), "''", "S") & " "		
	End If

    UNIValue(0,1) = strVal   
    
    UNIValue(1,0) = FilterVar(Trim(Request("txtItem")), "" , "S")  
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
       
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
	
	Set lgADF   = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
       
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
		
		If Len(Request("txtItem")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
		End If	
    Else    
	
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    	
    If BlankchkFlg = False Then
		If rs0.EOF And rs0.BOF Then
		   Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		   rs0.Close
		   Set rs0 = Nothing
		   Exit Sub
		Else    
			Call  MakeSpreadSheetData()	    
		End If
    End If


End Sub

%>
<Script Language=vbscript>

    parent.frm1.txtItemNm.Value	= "<%=ConvSPChars(arrRsVal(1))%>"
	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			parent.frm1.txtHFromDt.value	= "<%=Request("txtFromDt")%>"
			parent.frm1.txtHToDt.value		= "<%=Request("txtToDt")%>"
			parent.frm1.txtHItem.value	= "<%=ConvSPChars(Request("txtItem"))%>"
       End If
       'Show multi spreadsheet data from this line
       parent.ggoSpread.Source  = parent.frm1.vspdData
       parent.frm1.vspdData.Redraw = False
	   parent.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"
	   
	   Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,"<%=Request("txtCurrency")%>",Parent.GetKeyPos("A",11),"C", "Q" ,"X","X")		
	   Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,"<%=Request("txtCurrency")%>",Parent.GetKeyPos("A",12),"A", "Q" ,"X","X")		
	         
       
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
       parent.DbQueryOk
       parent.frm1.vspdData.Redraw = True
    End If   
</Script>	
