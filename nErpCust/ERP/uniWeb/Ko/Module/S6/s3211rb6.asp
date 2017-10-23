<%'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : B/L관리 
'*  3. Program ID           : s3211rb6.asp
'*  4. Program Name         : C/C 등록을 위한 L/C참조 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/10
'*  8. Modified date(Last)  : 2002/04/27
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : Kwak Eunkyoung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'*							: -2000/04/10 : Coding Start
'*                          : -2002/04/27 : ADO변환 
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2	  
Dim lgStrData                                                 
                                             
Dim lgTailList                                                
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim arrRsVal(3)
Dim BlankchkFlg

Dim iFrPoint
iFrPoint=0

    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "RB")
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    	   
	lgSelectList   = Request("lgSelectList")                               
	lgTailList     = Request("lgTailList")                                 	
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             
	lgDataExist      = "No"
	Const C_SHEETMAXROWS_D  = 30            
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
Sub SetConditionData()
End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,2)

    UNISqlId(0) = "S3211RA601"  										' main query(spread sheet에 뿌려지는 query statement)
	UNISqlId(1) = "s0000qa002"  										' 거래처코드/명 
	UNISqlId(2) = "s0000qa005"  										' 영업그룹코드/명 

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    

	strVal = " "
		

	IF Len(Trim(Request("txtApplicant"))) THEN
		strVal = strVal & " AND D.Applicant = " & FilterVar(Request("txtApplicant"), "''", "S") & " " 
		arrVal(0) = Trim(Request("txtApplicant"))
	END If
	
	
	IF Len(Trim(Request("txtSalesGroup"))) THEN 
		strVal = strVal & " AND D.SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & "  "
		arrVal(1) = Trim(Request("txtSalesGroup")) 
	END IF
	
	IF Len(Trim(Request("txtCurrency"))) THEN 
		strVal = strVal & " AND D.CUR = " & FilterVar(Trim(Request("txtCurrency")), "" , "S") & "  " 
	END IF
	IF Len(Trim(Request("txtDocAmt"))) THEN 
		strVal = strVal & " AND D.LC_AMT >= " & UniconvNum(Trim(Request("txtDocAmt")),0) & " " 
	END IF
	
	IF Len(Trim(Request("txtFromDt"))) THEN 
		strVal = strVal & " AND D.OPEN_DT >= " & FilterVar(UniconvDate(Trim(Request("txtFromDt"))), "''", "S") & " " 
	END IF

	IF Len(Trim(Request("txtToDt"))) THEN 
		strVal = strVal & " AND D.OPEN_DT <= " & FilterVar(UniconvDate(Trim(Request("txtToDt"))), "''", "S") & " " 
	END IF
     
    UNIValue(0,1) = strVal    												'	UNISqlId(0)의 두번째 ?에 입력됨	
	UNIValue(1,0) = FilterVar(Trim(Request("txtApplicant")), " " , "S")     						'☜: 거래처코드 
	UNIValue(2,0) = FilterVar(Trim(Request("txtSalesGroup")), " " , "S")    						'☜: 영업그룹코드 
    
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)


    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtApplicant")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수입처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
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
		.frm1.txtSalesGroupNm.value = "<%=ConvSPChars(arrRsVal(3))%>" 
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				parent.frm1.txthApplicant.value		= "<%=ConvSPChars(Request("txtApplicant"))%>"
				parent.frm1.txthSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>"
				parent.frm1.txthCurrency.value		= "<%=ConvSPChars(Request("txtCurrency"))%>"
				parent.frm1.txthDocAmt.value		= "<%=Request("txtDocAmt")%>"
				parent.frm1.txthFromDt.value		= "<%=Request("txtFromDt")%>"
				parent.frm1.txthToDt.value			= "<%=Request("txtToDt")%>"
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"
			
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")		
			
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			.DbQueryOk
			
			.frm1.vspdData.Redraw = True
		End If
	End with
</Script>	
