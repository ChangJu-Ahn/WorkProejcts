<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3112rb6.asp																*
'*  4. Program Name         : 수주내역참조(통관내역등록에서)											*
'*  5. Program Desc         : 수주내역참조(통관내역등록에서)											*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/07																*
'*  8. Modified date(Last)  : 2002/05/08																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Seo Jinkyung																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/07 : 화면 design												*
'*                            2. 2002/05/08 : Ado 변환													*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1          
	Dim lgStrData                                               
	Dim lgTailList
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo 
	Dim strVal
	Dim arrRsVal(1)
	Dim BlankchkFlg

	On Error Resume Next
	Err.Clear   

	Dim iFrPoint
	iFrPoint=0
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
    Call HideStatusWnd 
     
    lgPageNo         = UNICInt(Trim(Request("txtHlgPageNo")),0)              
    lgSelectList     = Request("txtHlgSelectList")
    lgTailList       = Request("txtHlgTailList")
    lgSelectListDT   = Split(Request("txtHlgSelectListDT"), gColSep)         
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                
    call  QueryData()                                                    
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    Const C_SHEETMAXROWS_D = 30
    
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

    If iLoopCount < C_SHEETMAXROWS_D Then                                      
       lgPageNo = ""
    End If
    rs0.Close                                                       
    Set rs0 = Nothing	                                            

End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim arrVal(0)
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(1,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                
	Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
				
	lgStrSQL = "SELECT REFERENCE FROM B_CONFIGURATION " 
	lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & FilterVar("S0017", "''", "S") & " AND MINOR_CD = " & FilterVar("A", "''", "S") & "  AND REFERENCE= " & FilterVar("Y", "''", "S") & " "
		
	IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then		
		UNISqlId(0) = "S3112RA601" 
	Else
		UNISqlId(0) = "S3112RA602" 
	End if
	Call SubCloseRs(lgObjRs)  
	Call SubCloseDB(lgObjConn)
     
     UNISqlId(1) = "s0000qa001"     
     
     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	strVal = " "

	If Len(Request("txtHSONo")) Then
		strVal = " AND SSD.so_no = " & FilterVar(Request("txtHSONo"), "''", "S") & " "	
	End If
	
	If Len(Request("txtHApplicant")) Then
		strVal = strVal & " AND SSH.sold_to_party = " & FilterVar(Request("txtHApplicant"), "''", "S") & " "	

	End If
	
	If Len(Request("txtHItem")) Then
		strVal = strVal & " AND BI.item_cd = " & FilterVar(Request("txtHItem"), "''", "S") & " "		
		arrVal(0) = Trim(Request("txtHItem"))
	End If

	If Len(Request("txtHCurrency")) Then
		strVal = strVal & " AND SSH.cur = " & FilterVar(Request("txtHCurrency"), "''", "S") & " "	
	End If
	
	If Len(Request("txtHSalesGroup")) Then
		strVal = strVal & " AND SSH.sales_grp  = " & FilterVar(Request("txtHSalesGroup"), "''", "S") & " "			
	End If
	
	If Len(Request("txtHPayTerms")) Then
		strVal = strVal & " AND SSH.pay_meth  = " & FilterVar(Request("txtHPayTerms"), "''", "S") & " "	
	End If
	
	If Len(Request("txtHIncoTerms")) Then
		strVal = strVal & " AND SSH.incoterms  = " & FilterVar(Request("txtHIncoTerms"), "''", "S") & " "			
	End If
		
    If Len(Trim(Request("txtHFromDt"))) Then
		strVal = strVal & " and SSH.so_dt >= " & FilterVar(UNIConvDate(Request("txtHFromDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtHToDt"))) Then
		strVal = strVal & " and SSH.so_dt <= " & FilterVar(UNIConvDate(Request("txtHToDt")), "''", "S") & ""		
	End If

	If Len(Trim(Request("txtHTrackingNo"))) Then
		strVal = strVal & " and  SSD.tracking_no = " & FilterVar(Request("txtHTrackingNo"), "''", "S") & " "		
	End If

    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(Trim(Request("txtHItem")), " " , "S")      
   
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

		If Len(Request("txtHItem")) And BlankchkFlg = False Then
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
      
       parent.ggoSpread.Source  = parent.frm1.vspdData
       
       parent.frm1.vspdData.Redraw = False
	   parent.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"
	
	   Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,"<%=Request("txtHCurrency")%>",Parent.GetKeyPos("A",9),"C", "Q" ,"X","X")		
	   Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,"<%=Request("txtHCurrency")%>",Parent.GetKeyPos("A",10),"A", "Q" ,"X","X")		
	   			
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
       parent.DbQueryOk
       parent.frm1.vspdData.Redraw = True
    End If   
</Script>	
