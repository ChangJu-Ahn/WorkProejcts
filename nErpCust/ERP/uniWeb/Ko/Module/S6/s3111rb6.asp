<%
'**********************************************************************************************
'*  1. Module Name          : 영업관리 
'*  2. Function Name        : 
'*  3. Program ID           : s3111rb6
'*  4. Program Name         : 수주참조(통관등록에서)
'*  5. Program Desc         : 수주참조(통관등록에서)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/02
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Choinkuk		
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3 			   
Dim lgStrData                                                 
Dim lgTailList                                                
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim arrRsVal(5)
Dim BlankchkFlg
Dim iFrPoint
iFrPoint=0
Const C_SHEETMAXROWS_D  = 30           
                               
	On Error Resume Next		
	Err.Clear
    
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "RB")
    Call HideStatusWnd     
    	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    	
	lgSelectList   = Request("lgSelectList")                               
	lgTailList     = Request("lgTailList")                                 
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             
	lgDataExist      = "No"
	    
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
    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,2)

	Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
				
	lgStrSQL = "SELECT REFERENCE FROM B_CONFIGURATION " 
	lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & FilterVar("S0017", "''", "S") & " AND MINOR_CD = " & FilterVar("A", "''", "S") & "  AND REFERENCE= " & FilterVar("Y", "''", "S") & " "
		
	IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then		
		UNISqlId(0) = "S3111RA601"
	Else
		UNISqlId(0) = "S3111RA602"
	End if
	Call SubCloseRs(lgObjRs)  
	Call SubCloseDB(lgObjConn)
    
    UNISqlId(1) = "s0000qa002"					'수입자명 
    UNISqlId(2) = "s0000qa005"					'영업그룹명 
    UNISqlId(3) = "s0000qa007"					'수주형태명  

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
	strVal = ""

	If Len(Request("txtApplicant")) Then
		strVal = strVal & "AND a.BP_CD =  " & FilterVar(Request("txtApplicant"), "''", "S") & "  "	
		arrVal(0) = Trim(Request("txtApplicant")) 
	End If

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & "AND b.SALES_GRP =  " & FilterVar(Request("txtSalesGroup"), "''", "S") & "  "		
		arrVal(1) = Trim(Request("txtSalesGroup")) 
	End If		
		
 	If Len(Request("txtSOType")) Then
		strVal = strVal & "AND d.SO_TYPE =  " & FilterVar(Request("txtSOType"), "''", "S") & "  "		
		arrVal(2) = Trim(Request("txtSOType")) 
	End If	    
	
    If Len(Request("txtFromDt")) Then
		strVal = strVal & "AND c.SO_DT >=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & " "			
	End If		
		
	If Len(Request("txtToDt")) Then
		strVal = strVal & "AND c.SO_DT <=  " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & " "		
	End If
	
    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(Trim(Request("txtApplicant")), " " , "S")					'수입자코드 
    UNIValue(2,0) = FilterVar(Trim(Request("txtSalesGroup")), " " , "S")				    '영업그룹코드 
    UNIValue(3,0) = FilterVar(Trim(Request("txtSOType")), " " , "S")					'수주형태코드    
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    BlankchkFlg = False
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2,rs3 )

    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtApplicant")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수입자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
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

		If Len(Request("txtSOType")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수주형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
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
		   Exit Sub
		Else    
		    Call  MakeSpreadSheetData()	    
		End If
    End If
	

End Sub

%>

<Script Language=vbscript>
    With parent
		.frm1.txtApplicantNm.value	= "<%=ConvSPChars(arrRsVal(1))%>" 
		.frm1.txtSalesGroupNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtSOTypeNm.value		= "<%=ConvSPChars(arrRsVal(5))%>"
        
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHApplicant.value	= "<%=ConvSPChars(Request("txtApplicant"))%>"
				.frm1.txtHSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>" 
				.frm1.txtHSOType.value		= "<%=ConvSPChars(Request("txtSOType"))%>"
				.frm1.txtHFromDt.value		= "<%=Request("txtFromDt")%>"
				.frm1.txtHToDt.value		= "<%=Request("txtToDt")%>"
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"
			
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.GetKeyPos("A",4),Parent.GetKeyPos("A",5),"A", "Q" ,"X","X")		
									
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			.DbQueryOk
			
			.frm1.vspdData.Redraw = True
			
		End If
	End with
</Script>	
