<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 입고번호 Popup
'*  3. Program ID           : M4111PB1
'*  4. Program Name         : 
'*  5. Program Desc         : 구매입고등록화면의 입고번호 팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2003/05/27
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :	
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/04 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/17 : ADO변환 
'**************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%							                           '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
	'On Error Resume Next													   '실행 오류가 발생할 때 오류가 발생한 문장 바로 다음에 실행이 계속될 수 있는 문으로 컨트롤을 옮길 수 있도록 지정합니다.				
	
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
	Dim rs1, rs2, rs3, rs4
	Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
	Dim iTotstrData
	Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo

	Dim strBeneficiaryNm
	Dim strPurGrpNm
	Dim strMvmtTypenm	
	
	Dim iPrevEndRow
    Dim iEndRow
	
	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")
	

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
	
	iPrevEndRow = 0
    iEndRow = 0
    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
	Dim strEnd
	
	Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(3,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 
	
	
    UNISqlId(0) = "M4111PA301"											' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(1) = "s0000qa024"	'공급처 
    UNISqlId(2) = "s0000qa019"	'구매그룹 
    UNISqlId(3) = "M4111PA302"	'입고유형(STO건으로 인한 수정(KJH:03-01-06)
    
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
																		  '	UNISqlId(0)의 첫번째 ?에 입력됨				
    strEnd   = "ZZZZZZZZZZZZZZZZZZ"
	
	strVal = " "

	If Len(Request("txtFrRcptDt")) Then 
		strVal = strVal & " AND A.MVMT_RCPT_DT >=  " & FilterVar(UNIConvDate(Request("txtFrRcptDt")), "''", "S") & " "
	End If		
	
	If Len(Request("txtToRcptDt")) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <=  " & FilterVar(UNIConvDate(Request("txtToRcptDt")), "''", "S") & " "
	End If
			
	
	If Len(Request("txtMvmtType")) Then
	  strVal = strVal & " AND B.IO_TYPE_CD =  " & FilterVar(Request("txtMvmtType"), "''", "S") & " "
	ELSE	
	   strVal = strVal & " AND B.IO_TYPE_CD <=  " & FilterVar(strEnd , "''", "S") & " "
	end if	    
    		
	If Len(Trim(Request("txtSupplier"))) Then
		strVal = strVal & " AND A.BP_CD =  " & FilterVar(Request("txtSupplier"), "''", "S") & " "
	ELSE	
		strVal = strVal & " AND A.BP_CD <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "
	End If
	
	IF LEN(Trim(Request("txtGroup"))) THEN
		strVal = strVal & " AND A.PUR_GRP =  " & FilterVar(Request("txtGroup"), "''", "S") & " "
	ELSE
	    strVal = strVal & " AND A.PUR_GRP <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "  
	END IF	 
 
    strVal = strVal & " AND (A.DLVY_ORD_FLG <> " & FilterVar("Y", "''", "S") & "  OR A.DLVY_ORD_FLG is Null) "
	
	UNIValue(0,0) = lgSelectList                                    '☜: Select list
	UNIValue(0,1) = strVal											'	UNISqlId(0)의 두번째 ?에 입력됨	
'	call svrmsgbox(strVal, vbinformation, i_mkscript)
	UNIValue(1,0) = FilterVar(Trim(Request("txtSupplier"))," ","S") 			    	'공급처 
	UNIValue(1,1) = " AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "											'사외거래처만 
	UNIValue(2,0) = FilterVar(Trim(Request("txtGroup"))," ","S") 						'구매그룹 
    UNIValue(3,0) = " AND A.IO_TYPE_CD =  " & FilterVar(Request("txtMvmtType"), "''", "S") & "  "	'입고유형 
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
  
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '☜: set ADO read mode
	
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim FalsechkFlg
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    If SetConditionData = false then Exit Sub
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
   
End Sub
    
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim PvArr
    
    Const C_SHEETMAXROWS_D = 100     
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                 
	End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
            PvArr(iLoopCount) = lgstrData	
		    lgstrData = ""
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        iEndRow = iPrevEndRow + iLoopCount + 1
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
   
    SetConditionData = false
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strBeneficiaryNm = rs1("Bp_Nm")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtSupplier"))) Then
			Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If 
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strPurGrpNm = rs2("Pur_Grp_Nm")
   		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	
	If Not(rs3.EOF Or rs3.BOF) Then
        strMvmtTypenm = rs3(1)
   		Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Trim(Request("txtMvmtType"))) Then
			Call DisplayMsgBox("970000", vbInformation, "입고형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	
	
	SetConditionData = true
	
End Function

%>

<Script Language=vbscript>
    With parent
		.frm1.txtSupplierNm.value = "<%=ConvSPChars(strBeneficiaryNm)%>" 
		.frm1.txtGroupNm.value = "<%=ConvSPChars(strPurGrpNm)%>" 
      	.frm1.txtMvmtTypeNm.value = "<%=ConvSPChars(strMvmtTypenm)%>" 

      	If "<%=lgDataExist%>" = "Yes" Then
			
			If "<%=lgPageNo%>" = "1" Then  
			    .frm1.hdnMvmtType.Value	 	= "<%=ConvSPChars(Request("txtMvmtType"))%>"				
				.frm1.hdnSupplier.Value 	= "<%=ConvSPChars(Request("txtSupplier"))%>"
				.frm1.hdnFrRcptDt.Value 	= "<%=Request("txtFrRcptDt")%>"
				.frm1.hdnToRcptDt.Value 	= "<%=Request("txtToRcptDt")%>"
				.frm1.hdnInspFlag.value     = "<%=ConvSPChars(Request("txtInspFlag"))%>"
				.frm1.hdnGroup.Value 	    = "<%=ConvSPChars(Request("txtGroup"))%>" 				
			End If    
			
			.ggoSpread.Source    = .frm1.vspdData 
			.frm1.vspdData.Redraw = false
		    .ggoSpread.SSShowData "<%=iTotstrData%>"                            '☜: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = True
		End If
	End with
</Script>	
 	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>

