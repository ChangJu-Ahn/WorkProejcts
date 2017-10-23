<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : M9211RB1													*
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2002/05/07																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : KO MYOUNG JIN
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"										*
'*                            this mark(☆) Means that "must change"										*
'* 13. History              : 1. 2000/04/08 : Coding Start												*
'********************************************************************************************************


'On Error Resume Next
   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2           '☜ : DBAgent Parameter 선언 
   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
   Dim strShiptoPartyNm
   Dim strPlantNm
   Dim strSlNm
	   
   Dim lgF0,i
   Dim iCodeArr
   Dim strPurGrpNm
   DIM strBP_NM
'--------------- 개발자 coding part(변수선언,Start)----------------------------------------------------


'--------------- 개발자 coding part(변수선언,End)------------------------------------------------------

	Call HideStatusWnd
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = CInt(Request("lgMaxCount"))                       '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgStrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgStrData = lgStrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	

    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    
    SetConditionData = FALSE

    
    If Not(rs1.EOF Or rs1.BOF) Then
        strPurGrpNm = rs1("Pur_Grp_Nm")
    
   		Set rs1 = Nothing
    Else

		Set rs1 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			exit function
		End If
	End If 
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strBP_NM = rs2("BP_NM")
    
   		Set rs2 = Nothing
    Else

		Set rs2 = Nothing
		If Len(Trim(Request("txtSGICd"))) Then
			Call DisplayMsgBox("970000", vbInformation, "출고공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			exit function
		End If
	End If 
	
	
	SetConditionData = TRUE
	
End Function


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,2)
	
    UNISqlId(0) = "M9211RA101"									'* : 데이터 조회를 위한 SQL문     
    UNISqlId(1) = "S0000QA022"	'구매그룹 
	'UNISqlId(2) = "M9211BPCD"
	UNISqlId(2) = "Q2111QA123"
	
    UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	If Len(Request("txtStoNo")) Then
		strVal = strVal & " AND A.CUST_PO_NO = " & FilterVar(Request("txtStoNo"), "''", "S") & " "
	End If

    If Len(Trim(Request("txtFrStoDt"))) Then
		strVal = strVal & " AND A.SO_DT >= " & FilterVar(UNIConvDate(Request("txtFrStoDt")), "''", "S") & ""
	End If		
	
	If Len(Trim(Request("txtToStoDt"))) Then
		strVal = strVal & " AND A.SO_DT <= " & FilterVar(UNIConvDate(Request("txtToStoDt")), "''", "S") & ""		
	End If

	'If Len(Request("txtSGINo")) Then
	'	strVal = strVal &  " AND C.DN_NO ='" & FilterVar(Trim(Request("txtSGINo")), " " , "SNM") & "'"
		
	'End If

    If Len(Trim(Request("txtFrStoDt"))) Then
		strVal = strVal & " AND C1.DLVY_DT >= " & FilterVar(UNIConvDate(Request("txtFrSGIDt")), "''", "S") & ""
	End If		
	
	If Len(Trim(Request("txtToStoDt"))) Then
		strVal = strVal & " AND C1.DLVY_DT <= " & FilterVar(UNIConvDate(Request("txtToSGIDt")), "''", "S") & ""		
	End If
	
	If Len(Request("txtSGICd")) Then
	'strVal = strVal & " AND H.BP_CD = 'SD001'"
		strVal = strVal & " AND H.BP_CD = " & FilterVar(Request("txtSGICd"), "''", "S") & " "
	End If
	
	If Len(Request("txtGroup")) Then
	'strVal = strVal & " AND H.BP_CD = 'SD001'"
		strVal = strVal & " AND H.PUR_GRP = " & FilterVar(Request("txtGroup"), "''", "S") & " "
	End If


	strVal = strVal & " Order by g.po_no, g.po_seq_no desc"
    
   
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

    UNIValue(0,1) = strVal       
        
    UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S")					'구매그룹 
    UNIValue(2,0) = " " & FilterVar(UCase(Request("txtSGICd")), "''", "S") & " "					'구매그룹 
'================================================================================================================   
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
	IF SetConditionData() = FALSE THEN EXIT SUB
 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub


%>
<Script Language=vbscript>
	
	parent.frm1.txtGroupNm.value		= "<%=ConvSPChars(strPurGrpNm)%>" 
	parent.frm1.txtSGiNm.value			= "<%=ConvSPChars(strBP_NM)%>"
    If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			parent.frm1.hdnStoNo.value		= "<%=ConvSPChars(Request("txtStoNO"))%>"   
			parent.frm1.hdnFrStoDt.value		= "<%=Request("txtFrStoDt")%>"
			parent.frm1.hdnToStoDt.value		= "<%=Request("txtToStoDt")%>"
			'parent.frm1.hdnSGINo.value		= "<%=ConvSPChars(Request("txtSGINO"))%>"
			parent.frm1.hdnSGICd.value		= "<%=ConvSPChars(Request("txtSGICd"))%>"      
			parent.frm1.hdnFrSGIDt.value		= "<%=Request("txtFrSGIDt")%>"
			parent.frm1.hdnToSGIDt.value		= "<%=Request("txtToSGIDt")%>"
       End If
       'Show multi spreadsheet data from this line
       
       parent.ggoSpread.Source  = parent.frm1.vspdData
       parent.ggoSpread.SSShowData "<%=lgstrData%>"          '☜ : Display data
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       parent.DbQueryOk
    End If   
</Script>	
