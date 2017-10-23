<!--
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : L/C 관리 
'*  3. Program ID           : M3211PB1
'*  4. Program Name         : Open L/C No POPUP ASP
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2002/04/16
'*  9. Modifier (First)     : Min, HJ
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/04 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/16 : ADO변환 
'**************************************************************************************
%>
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next													   '실행 오류가 발생할 때 오류가 발생한 문장 바로 다음에 실행이 계속될 수 있는 문으로 컨트롤을 옮길 수 있도록 지정합니다.				

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1       	'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgpageNo	                                            '☜ : 이전 값 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim iTotstrData

Dim strBeneficiaryNm

Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)              '☜ : Next key flag
lgSelectList    = Request("lgSelectList")
lgTailList      = Request("lgTailList")
lgSelectListDT  = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
lgDataExist     = "No"							                    '☜ : Orderby value

Call FixUNISQLData()
Call QueryData()


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
	Dim strEnd
	
	Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(1,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 
	
	strEnd   = "ZZZZZZZZZZZZZZZZZZ"
	
    UNISqlId(0) = "m3211pa101"												' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(1) = "s0000qa024"												' 2번째 query(spread sheet에 뿌려지는 query statement)
    
    strVal = " WHERE A.LC_KIND = " & FilterVar("M", "''", "S") & "   "
	
	If Len(Trim(Request("txtFrDt"))) Then
		strVal = strVal & " AND A.REQ_DT >= " & FilterVar(UNIConvDate(Request("txtFrDt")), "''", "S") & " "
	End If

	If Len(Trim(Request("txtToDt"))) Then
		strVal = strVal & " AND A.REQ_DT <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & " "
	End If
		
	
	If Len(Request("txtLCNo")) Then
		strVal = strVal & " AND A.LC_NO =  " & FilterVar(Trim(UCase(Request("txtLCNo"))), " " , "S") & " "
	else
		strVal = strVal & " AND A.LC_NO <=  " & FilterVar(strEnd , "''", "S") & " "
	end if
	
	If Len(Trim(Request("txtBeneficiary"))) Then
		strVal = strVal & " AND A.BENEFICIARY =  " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " "
	else
		strVal = strVal & " AND A.BENEFICIARY <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "
	End If
	
	If Len(Trim(Request("gBizArea"))) Then
		strVal = strVal & " AND A.BIZ_AREA =  " & FilterVar(Request("gBizArea"), "''", "S") & " "
	End If
	If Len(Trim(Request("gPurGrp"))) Then
		strVal = strVal & " AND A.PUR_GRP =  " & FilterVar(Request("gPurGrp"), "''", "S") & " "
	End If
	If Len(Trim(Request("gPurOrg"))) Then
		strVal = strVal & " AND A.PUR_ORG =  " & FilterVar(Request("gPurOrg"), "''", "S") & " "
	End If
		   
	UNIValue(0,0) = lgSelectList                                          '☜: Select list
	UNIValue(0,1) = strVal												'	UNISqlId(0)의 두번째 ?에 입력됨	
	UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S")				  		  '	UNISqlId(1)의 첫번째 ?에 입력됨		
	
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    'UNIValue(0,UBound(UNIValue,2)    ) = " ORDER BY A.LC_NO DESC "
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Set lgADF   = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    IF SetConditionData() = FALSE THEN EXIT SUB
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
   
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
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
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
  	If CLng(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
           PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")

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
    On Error Resume Next
    
    SetConditionData = FALSE
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strBeneficiaryNm = rs1("BP_NM")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtBeneficiary"))) Then
			Call DisplayMsgBox("970000", vbInformation, "수출자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			exit function
		End If
		
	End If   	
    SetConditionData = TRUE 
End Function    
											'☜: 비지니스 로직 처리를 종료함 
%>
<Script Language=vbscript>
    With parent

		.frm1.txtBeneficiaryNm.value = "<%=ConvSPChars(strBeneficiaryNm)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			
			If "<%=lgPageNo%>" = "1" Then          
				.frm1.hdnBeneficiary.value	= "<%=ConvSPChars(Request("txtBeneficiary"))%>"
				.frm1.hdnFrDt.value			= "<%=Request("txtFrDt")%>"
				.frm1.hdnToDt.value			= "<%=Request("txtToDt")%>"
			End If    
			
			.ggoSpread.Source = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"					'☜: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
