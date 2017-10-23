<%'
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4111ra6.asp																*
'*  4. Program Name         : 외주출고참조(통관등록에서)												*
'*  5. Program Desc         : 외주출고참조(통관등록에서)												*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2002/04/11																*
'*  8. Modified date(Last)  : 2002/07/10																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son Bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
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
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,2)
	
	On Error Resume Next
	
    UNISqlId(0) = "M4111RA601"
    UNISqlId(1) = "s0000qa002"					'수입자명 
    UNISqlId(2) = "s0000qa019"					'구매그룹명 
    UNISqlId(3) = "s0000qa009"					'공장명  
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = ""
	
	
	
	If Len(Request("txtApplicant")) Then
		strVal = strVal & "AND A.BP_CD =  " & FilterVar(Request("txtApplicant"), "''", "S") & "  "	
		arrVal(0) = Trim(Request("txtApplicant")) 
	End If

	If Len(Request("txtPurGroup")) Then
		strVal = strVal & "AND B.PUR_GRP =  " & FilterVar(Request("txtPurGroup"), "''", "S") & "  "		
		arrVal(1) = Trim(Request("txtPurGroup")) 
	End If		
		
 	If Len(Request("txtPlant")) Then
		strVal = strVal & "AND E.PLANT_CD =  " & FilterVar(Request("txtPlant"), "''", "S") & "  "		
		arrVal(2) = Trim(Request("txtPlant")) 
	End If	    
	
    If Len(Request("txtFromDt")) Then
		strVal = strVal & "AND C.PO_DT >=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & " "			
	End If		
	
	If Len(Request("txtToDt")) Then
		strVal = strVal & "AND C.PO_DT <=  " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & " "		
	End If
	
    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(Trim(Request("txtApplicant")), " " , "S")			'수입자코드 
    UNIValue(2,0) = FilterVar(Trim(Request("txtPurGroup")), " " , "S")		    '영업그룹코드 
    UNIValue(3,0) = FilterVar(Trim(Request("txtPlant")), " " , "S")			'수주형태코드    
    
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    BlankchkFlg = False
    On Error Resume Next
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

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

		If Len(Request("txtPurGroup")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
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

		If Len(Request("txtPlant")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
		End If	
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs1.Close
        Set rs3 = Nothing
    End If
	

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 


'    If BlankchkFlg = False Then         
	If  rs0.EOF And rs0.BOF And BlankchkFlg = False Then
		rs0.Close
	    Set rs0 = Nothing
	    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	Else    
	    Call  MakeSpreadSheetData()		    
	End If  
'	End If	
End Sub

%>

<Script Language=vbscript>
    With parent
		.frm1.txtApplicantNm.value	= "<%=ConvSPChars(arrRsVal(1))%>" 
		.frm1.txtPurGroupNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtPlantNm.value		= "<%=ConvSPChars(arrRsVal(5))%>"
        
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHApplicant.value	= "<%=ConvSPChars(Request("txtApplicant"))%>"
				.frm1.txtHPurGroup.value	= "<%=ConvSPChars(Request("txtPurGroup"))%>" 
				.frm1.txtHPlant.value		= "<%=ConvSPChars(Request("txtPlant"))%>"
				.frm1.txtHFromDt.value		= "<%=Request("txtFromDt")%>"
				.frm1.txtHToDt.value		= "<%=Request("txtToDt")%>"
			End If    
			'Show multi spreadsheet data from this line
			       
			.frm1.vspdData.Redraw = False
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '☜: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			.frm1.vspdData.Redraw = True
			.DbQueryOk
		End If
	End with
</Script>	


