<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M3111PB7
'*  4. Program Name         : 반품발주번호 
'*  5. Program Desc         : 반품발주번호 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Min, HJ
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
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1, rs2, rs3	'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgpageNo	                                            '☜ : 이전 값 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim iFrPoint
iFrPoint=0

'--------------- 개발자 coding part(변수선언,Start)----------------------------------------------------
Dim strPotypeNm, strSupplierNm, strGroupNm

Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")
     
lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)              '☜ : Next key flag
lgSelectList    = Request("lgSelectList")
lgTailList      = Request("lgTailList")
lgSelectListDT  = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
lgDataExist     = "No"
	 
Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strVal
	Dim sTemp
	
	Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(4,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
    
    '--- 2004-08-19 by Byun Jee Hyun for UNICODE                                                                      '    parameter의 수에 따라 변경함 
	strVal = ""
    UNISqlId(0) = "M3111PA701"
    UNISqlId(1) = "s0000qa020"	'발주형태 
    UNISqlId(2) = "s0000qa024"	'공급처 
    UNISqlId(3) = "s0000qa022"	'구매그룹 
    
	'발주형태                    
    If Len(Trim(Request("txtPotypeCd"))) Then
		strVal = strVal & " AND A.PO_TYPE_CD =  " & FilterVar(Trim(UCase(Request("txtPotypeCd"))), " " , "S") & "  "	
	End If
    '확정여부			
    If Trim(Request("txtRadio")) = "Y" then
		strVal = strVal & " AND A.RELEASE_FLG = " & FilterVar("Y", "''", "S") & "  "
	ElseIf Trim(Request("txtRadio")) = "N" then
		strVal = strVal & " AND A.RELEASE_FLG = " & FilterVar("N", "''", "S") & "  "
	End If
	'공급처 
    If Len(Trim(Request("txtSupplierCd"))) Then
		strVal = strVal & " AND A.BP_CD =  " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & "  "	
	End If
    '발주일 
    If Len(Trim(Request("txtFrPoDt"))) Then
		strVal = strVal & " AND A.PO_DT >=  " & FilterVar(UniConvDate(Request("txtFrPoDt")), "''", "S") & " "	
	End If
			
    If Len(Trim(Request("txtToPoDt"))) Then
		strVal = strVal & " AND A.PO_DT <=  " & FilterVar(UniConvDate(Request("txtToPoDt")), "''", "S") & " "	
	End If
	'구매구룹 
	If Len(Trim(Request("txtGroupCd"))) Then
		strVal = strVal & " AND A.PUR_GRP =  " & FilterVar(Trim(UCase(Request("txtGroupCd"))), " " , "S") & "  "	
	End If
    '반품 
	If Trim(Request("hdnRetFlg")) = "Y" then
		strVal = strVal & " AND A.RET_FLG = " & FilterVar("Y", "''", "S") & "  "
	ElseIf Trim(Request("hdnRetFlg")) = "N" then
		strVal = strVal & " AND A.RET_FLG = " & FilterVar("N", "''", "S") & "  "
	End If
	%>
	<SCRIPT Language="vBS">
	
	</SCRIPT>
	<%
	
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절 필드 
	UNIValue(0,1) = strVal													'---WHERE 절 
	 
	UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtPotypeCd"))), " " , "S")  		  '	UNISqlId(1)의 첫번째 ?에 입력됨	
	UNIValue(2,0) = FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S")   	  '	UNISqlId(2)의 첫번째 ?에 입력됨	
	UNIValue(3,0) = FilterVar(Trim(UCase(Request("txtGroupCd"))), " " , "S")  		  '	UNISqlId(3)의 첫번째 ?에 입력됨	
    
    UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)					'---Order By 조건 
    'UNIValue(0,UBound(UNIValue,2)    ) = " ORDER BY A.PO_NO DESC "
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
	Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim iStr
   
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF   = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
     
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
   
    IF SetConditionData() = FALSE THEN EXIT SUB
   
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    
    Else
		Call  MakeSpreadSheetData()        
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	On Error Resume Next
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
	Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
  	If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		iFrPoint    = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
	End If
    
    iLoopCount = -1
   
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
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
        
        PvArr(iLoopCount) = lgstrData
        lgstrData=""
        rs0.MoveNext
	Loop
    lgstrData = Join(PvArr,"")

    If iLoopCount < C_SHEETMAXROWS_D Then                                 '☜: Check if next data exists
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
     
    SetConditionData = FALSE
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strPotypeNm = rs1("PO_TYPE_NM")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtPotypeCd"))) Then
			Call DisplayMsgBox("970000", vbInformation, "반품형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
		
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strSupplierNm = rs2("BP_NM")
		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtSupplierCd"))) Then
			Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If			
    End If     
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strGroupNm = rs3("PUR_GRP_NM")
		Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Trim(Request("txtGroupCd"))) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If			
    End If  
     
    SetConditionData = TRUE
    
End Function


%>

<Script Language=vbscript>
    With parent

		.frm1.txtPotypeNm.value		= "<%=ConvSPChars(strPotypeNm)%>"
		.frm1.txtSupplierNm.value	= "<%=ConvSPChars(strSupplierNm)%>"
		.frm1.txtGroupNm.value		= "<%=ConvSPChars(strGroupNm)%>"

		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.hdnPotype.Value	 	= "<%=ConvSPChars(Request("txtPotypeCd"))%>"
				.frm1.hdnSupplier.Value 	= "<%=ConvSPChars(Request("txtSupplierCd"))%>"
				.frm1.hdnFrDt.Value 		= "<%=ConvSPChars(Request("txtFrPoDt"))%>"
				.frm1.hdnToDt.Value 		= "<%=ConvSPChars(Request("txtToPoDt"))%>"
				.frm1.hdnGroup.Value 		= "<%=ConvSPChars(Request("txtGroupCd"))%>"
				.frm1.hdnRadio.value		= "<%=ConvSPChars(Request("txtRadio"))%>"			
			End If    
			'Show multi spreadsheet data from this line
			.ggoSpread.Source = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=lgstrData%>","F"					'☜: Display data 

	        Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,Parent.GetKeyPos("A",8), Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")	'발주금액 

			.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
