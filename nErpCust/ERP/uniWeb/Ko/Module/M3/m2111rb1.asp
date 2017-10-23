<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m2111ra1
'*  4. Program Name         : 구매요청참조 
'*  5. Program Desc         : 구매요청참조 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/21	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Shin jin hyun		
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0       		   '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim strPtnBpNm												  ' 남품처명 
Dim strDNTypeNm												  ' 출하형태명 
Dim strSOTypeNm											      ' 수주타입명 

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
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
	Const C_SHEETMAXROWS_D  = 100            
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
    Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(0,2)

    UNISqlId(0) = "M2111RA102"
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtSoNo")) Then
		strVal = strVal & "AND A.SO_NO = " & FilterVar(Trim(UCase(Request("txtSoNo"))), " " , "S") & " "	
	End If

	If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), " " , "S") & "  "		
	End If		
		   
	If Len(Request("txtSupplier")) Then
		If Request("txtSTOflg") = "Y" then					
 			strVal = strVal & " AND G.SPPL_CD = '' "		'2002-12-16(LJT)
 		Else
 			strVal = strVal & " AND G.SPPL_CD =  " & FilterVar(Trim(UCase(Request("txtSupplier"))), " " , "S") & "  "		 			
 		End If
	End If	
    
 	If Len(Request("txtGroup")) Then
		strVal = strVal & " AND G.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S") & "  "		
	End If	    
	
 	If Len(Request("txtSubconfraflg")) Then
 		If Request("txtSubconfraflg") = "Y" AND Request("txtSTOflg") = "N" then
 			strVal = strVal & " AND A.PROCURE_TYPE <> " & FilterVar("P", "''", "S") & "  "		
		ElseIf Request("txtSubconfraflg") = "Y" AND Request("txtSTOflg") = "Y" then
 			strVal = strVal & " AND A.PROCURE_TYPE = " & FilterVar("P", "''", "S") & "  "		'2002-12-16(LJT)
 		Else
 			strVal = strVal & " AND A.PROCURE_TYPE = " & FilterVar("P", "''", "S") & "  "		
		End if		
	End If	      	
	
    If Len(Request("txtFrPoDt")) Then
		strVal = strVal & " AND G.PUR_PLAN_DT >= " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & " "			
	else
	    strVal = strVal & " AND G.PUR_PLAN_DT >= " & FilterVar("1900-01-01", "''", "S") & " "			
	End If		
	
	If Len(Request("txtToPoDt")) Then
		strVal = strVal & " AND G.PUR_PLAN_DT <= " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & " "		
	else
	    strVal = strVal & " AND G.PUR_PLAN_DT <= " & FilterVar("2999-12-31", "''", "S") & " "		
	End If
	
	If Len(Request("txtFrDlvyDt")) Then
		strVal = strVal & " AND A.DLVY_DT >= " & FilterVar(UNIConvDate(Request("txtFrDlvyDt")), "''", "S") & " "		
	End If		
	
	If Len(Request("txtToDlvyDt")) Then
		strVal = strVal & " AND A.DLVY_DT <= " & FilterVar(UNIConvDate(Request("txtToDlvyDt")), "''", "S") & " "		
	End If	

	If Len(Trim(Request("txtPlantCd"))) Then
		strVal = strVal & " AND A.PLANT_CD = " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "		
	End If	

    UNIValue(0,1) = strVal   & UCase(Trim(lgTailList)) 
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.hdnFrDt.Value 		= "<%=Request("txtFrPoDt")%>"
				.frm1.hdnToDt.Value 		= "<%=Request("txtToPoDt")%>"
				.frm1.hdnFrDt2.Value 		= "<%=Request("txtFrDlvyDt")%>"
				.frm1.hdnToDt2.Value 		= "<%=Request("txtToDlvyDt")%>"
				.frm1.hdnSoNo.value		= "<%=ConvSPChars(Request("txtSoNo"))%>"			
				.frm1.hdnTrackingNo.value	= "<%=ConvSPChars(Request("txtTrackingNo"))%>"			
			End If    
			'Show multi spreadsheet data from this line
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
