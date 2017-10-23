<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        :
'*  3. Program ID           : M2111QB1
'*  4. Program Name         : 구매요청조회 
'*  5. Program Desc         : 구매요청조회 
'*  6. Component List       :
'*  7. Modified date(First) : 2000/06/06
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : KANG SU HWAN
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
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	On Error Resume Next														'☜:
																'☆ : 입력/수정용 ComProxy Dll 사용 변수 

	Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
	Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
	Dim rs1, rs2, rs3, rs4, rs5, rs6							'☜ : DBAgent Parameter 선언 
	Dim lgstrData                                                              '☜ : data for spreadsheet data
	Dim lgStrPrevKey                                                           '☜ : 이전 값 
	Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT
	'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
	Dim ICount  		                                        '   Count for column index
	Dim strPlantCd                                              '   공장 
	Dim strPlantCdFrom
	Dim strItemCd												'   품목 
	Dim strItemCdFrom
	Dim strPrFrDt                                               '   요청일 
	Dim strPrToDt
	Dim strPdFrDt                                               '   필요일 
	Dim strPdToDt
	Dim strPrStsCd												'   요청진행상태 
	Dim strPrStsCdFrom
	Dim StrRqDeptCd												'	요청부서 
	Dim StrRqDeptCdFrom
	Dim strTrackNo

	Dim arrRsVal(11)											'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array
	Dim lgPageNo
	Dim lgDataExist
	'--------------- 개발자 coding part(변수선언,End)------------------------------------------------------
    Call HideStatusWnd
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 

    Call  TrimData()                                                     '☜ : Parent로 부터의 데이타 가공 
    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
    Const C_SHEETMAXROWS_D = 100
    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt
    Dim PvArr

    lgDataExist    = "Yes"
    lgstrData      = ""

    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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

	lgstrData  = Join(PvArr, "")

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

    Redim UNISqlId(6)                                                     '☜: SQL ID 저장을 위한 영역확보 

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(6,15)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "M2111QA101"
	 																	  '* : 각각의 조회조건부마다 Name 을 가져오는 SQL 문을 만듬 
     UNISqlId(1) = "M2111QA302"								              '공장명 
	 UNISqlId(2) = "M2111QA303"											  '품목명 
	 UNISqlId(3) = "M2111QA304"											  '(구매)요청진행상태명 
	 UNISqlId(4) = "M2111QA305"											  '(요청)부서명 
'	 UNISqlId(5) = "s0000qa017"										  '트레킹넘버 검색 
																		  'Reusage is Recommended
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

     UNIValue(0,0) = lgSelectList		                              '☜: Select 절에서 Summary    필드 

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     UNIValue(0,1)  = " " & FilterVar(Trim(UCase(Request("txtchangorgid"))), " " , "S") & " "
     UNIValue(0,2)  = UCase(Trim(strPlantCdFrom))		'---공장 
     UNIValue(0,3)  = UCase(Trim(strPlantCd))
     UNIValue(0,4)  = UCase(Trim(strItemCdFrom))		'---품목 
     UNIValue(0,5)  = UCase(Trim(strItemCd))
     UNIValue(0,6)  = UCase(Trim(strPrStsCdFrom))		'---요청진행상태 
     UNIValue(0,7)  = UCase(Trim(strPrStsCd))
     UNIValue(0,8)  = UCase(Trim(strPdFrDt))			'---필요일 
     UNIValue(0,9)  = UCase(Trim(strPdToDt))
     UNIValue(0,10) = UCase(Trim(strPrFrDt))			'---요청일 
     UNIValue(0,11) = UCase(Trim(strPrToDt))
     UNIValue(0,12) = UCase(Trim(strRqDeptCdFrom))	    '---요청부서 
     UNIValue(0,13) = UCase(Trim(strRqDeptCd))
     UNIValue(0,14) = UCase(Trim(strTrackNo))

	 UNIValue(1,0) = UCase(Trim(strPlantCd))
     UNIValue(2,0) = UCase(Trim(strPlantCd))
     UNIValue(2,1) = UCase(Trim(strItemCd))
     UNIValue(3,0) = UCase(Trim(strPrStsCd))
     UNIValue(4,0) = UCase(Trim(strRqDeptCd))
     'UNIValue(4,1) = UCase(Trim(Request("txtchangorgid")))
     UNIValue(4,1) = " " & FilterVar(Trim(UCase(Request("txtchangorgid"))), " " , "S") & " "
'      UNIValue(5,0) = UCase(Trim(strTrackNo))

'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

     UNIValue(0,UBound(UNIValue,2)) = " " & Trim(lgTailList)	'---Order By 조건 

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

    Dim FalsechkFlg

    FalsechkFlg = False

    '============================= 추가된 부분 =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
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
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
			Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			FalsechkFlg = True
			rs0.Close
			Set rs0 = Nothing
			Exit Sub		'20030124 - leejt
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
        If Len(Request("txtStateCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "요청진행상태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
		End If
    Else
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
        If Len(Request("txtDeptCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "요청부서", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
		End If
    Else
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If


'     If  rs5.EOF And rs5.BOF Then
'         rs5.Close
'         Set rs5 = Nothing
'         If Len(Request("txtTrackNo")) And FalsechkFlg = False Then
' 		   Call DisplayMsgBox("970000", vbInformation, "Tracking No", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
' 	       FalsechkFlg = True
' 		End If
'     Else
'         rs5.Close
'         Set rs5 = Nothing
'     End If

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else
        Call  MakeSpreadSheetData()
    End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    '---공장 
    If Len(Trim(Request("txtPlantCd"))) Then
    	strPlantCd	= " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
    	strPlantCdFrom = strPlantCd
    Else
    	strPlantCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPlantCdFrom = "''"
    End If

    '---품목 
    If Len(Trim(Request("txtItemCd"))) Then
    	strItemCd	= " " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
    	strItemCdFrom = strItemCd
    Else
    	strItemCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strItemCdFrom = "''"
    End If

	'---요청진행상태 
    If Len(Trim(Request("txtStateCd"))) Then
    	strPrStsCd	= " " & FilterVar(Trim(UCase(Request("txtStateCd"))), " " , "S") & " "
    	strPrStsCdFrom = strPrStsCd
    Else
    	strPrStsCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPrStsCdFrom = "''"
    End If

    '---필요일 
    If Len(Trim(Request("txtDlvyFrDt"))) Then
    	strPdFrDt 	= " " & FilterVar(uniConvDate(Request("txtDlvyFrDt")), "''", "S") & ""
    Else
    	strPdFrDt	= "" & FilterVar("1900/01/01", "''", "S") & ""
    End If

    If Len(Trim(Request("txtDlvyToDt"))) Then
    	strPdToDt 	= " " & FilterVar(uniConvDate(Request("txtDlvyToDt")), "''", "S") & ""
    Else
    	strPdToDt	= "" & FilterVar("2999/12/30", "''", "S") & ""
    End If

    '---요청일 
    If Len(Trim(Request("txtReqFrDt"))) Then
    	strPrFrDt 	= " " & FilterVar(uniConvDate(Request("txtReqFrDt")), "''", "S") & ""
    Else
    	strPrFrDt	= "" & FilterVar("1900/01/01", "''", "S") & ""
    End If

    If Len(Trim(Request("txtReqToDt"))) Then
    	strPrToDt 	= " " & FilterVar(uniConvDate(Request("txtReqToDt")), "''", "S") & ""
    Else
    	strPrToDt	= "" & FilterVar("2999/12/30", "''", "S") & ""
    End If

    '---요청부서 
    If Len(Trim(Request("txtDeptCd"))) Then
    	strRqDeptCd	= " " & FilterVar(Trim(UCase(Request("txtDeptCd"))), " " , "S") & " "
    	strRqDeptCdFrom = strRqDeptCd
    Else
    	strRqDeptCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strRqDeptCdFrom = "''"
    End If

    If Len(Trim(Request("txtTrackNo"))) Then
    	strTrackNo 	= " " & FilterVar(Trim(Request("txtTrackNo")), "''", "S") & ""
    Else
    	strTrackNo	= " A.TRACKING_NO "
    End If

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>

    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
         .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag

		.frm1.hdnDFrDt.Value = "<%=ConvSPChars(Request("txtDlvyFrDt"))%>"
		.frm1.hdnDToDt.Value = "<%=ConvSPChars(Request("txtDlvyToDt"))%>"
		.frm1.hdnRFrDt.Value = "<%=ConvSPChars(Request("txtReqFrDt"))%>"
		.frm1.hdnRToDt.Value = "<%=ConvSPChars(Request("txtReqToDt"))%>"
		.frm1.hdnPlant.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hdnItem.value  = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hdnState.Value = "<%=ConvSPChars(Request("txtStateCd"))%>"
		.frm1.hdnDept.value  = "<%=ConvSPChars(Request("txtDeptCd"))%>"
		.frm1.hdnTrackNo.value  = "<%=ConvSPChars(Request("txtTrackNo"))%>"

		.frm1.txtPlantNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>"
  		.frm1.txtItemNm.value			=  "<%=ConvSPChars(arrRsVal(3))%>"
  		.frm1.txtStateNm.value			=  "<%=ConvSPChars(arrRsVal(5))%>"
  		.frm1.txtDeptNm.value		=  "<%=ConvSPChars(arrRsVal(7))%>"
		.DbQueryOk(1)
	End With
</Script>
<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>
