<% Option Explicit %>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<% 

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("Q","B", "COOKIE", "QB")
      
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2 , rs3 , rs4        '☜ : DBAgent Parameter 선언
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim lgtxtAccountYear
Dim lgtxtBizArea
Dim lgtxtdeptcd
Dim lgtxtMaxRows

Dim biz_area_nm
Dim cost_nm
Dim dept_nm
Dim lgChangeOrgId
Dim strGlInputType,strGlInputTypeNM, StrDesc, strRefNo
Dim strWhere0
Dim strAmtFr, strAmtTo
Dim lgtxtCOST_CENTER_CD
Const C_SHEETMAXROWS_D  = 100                                  '☆: Server에서 한번에 fetch할 최대 데이타 건수

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call HideStatusWnd 

    lgPageNo       = Request("lgPageNo")                               '☜ : Next key flag
    lgMaxCount     = C_SHEETMAXROWS_D
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"

	lgtxtAccountYear	= Trim(Request("txtAccountYear"))
	lgtxtBizArea		= Trim(Request("txtBizArea"))
	lgtxtdeptcd			= Trim(Request("txtdeptcd"))
	lgtxtMaxRows		= Request("txtMaxRows")
	lgChangeOrgId		= Trim(Request("hChangeOrgId"))
	strGlInputType		= Trim(Request("txtGlInputType"))
	StrDesc		= Trim(Request("txtDesc"))
	strRefNo   = Trim(Request("txtRefNo"))
	Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'========================================================================================
' Query Data
'========================================================================================
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""


    If Len(Trim(lgPageNo))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          lgPageNo = CInt(lgPageNo)
       End If
    Else
       lgPageNo = 0
    End If

    'rs0에 대한 결과
    rs0.PageSize     = lgMaxCount                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1
    iLoopCount = -1

    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If

	rs0.Close
    Set rs0 = Nothing

    'rs1에 대한 결과
    IF NOT (rs1.EOF or rs1.BOF) then
	    biz_area_nm = rs1("biz_area_nm")
    END IF
    rs1.Close
    Set rs1 = Nothing

   'rs2에 대한 결과
    IF NOT (rs2.EOF or rs2.BOF) then
		dept_nm = rs2("dept_nm")
    END IF
    rs2.Close
    Set rs2 = Nothing

    'rs3에 대한 결과    
    IF NOT (rs3.EOF or rs3.BOF) then
		cost_nm = rs3("cost_nm")
    END IF
    rs3.Close
    Set rs3 = Nothing

    'rs3에 대한 결과    
    IF NOT (rs4.EOF or rs4.BOF) then
		strGlInputTypeNM= Trim(rs4("minor_nm"))
		strGlInputType	= Trim(rs4("minor_cd"))
    END IF
    rs4.Close
    Set rs4 = Nothing

End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(5)                                                    '☜: SQL ID 저장을 위한 영역확보
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(5,5)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수

    UNISqlId(0) = "a5131QA101"
    UNISqlId(1) = "ABIZNM"
    UNISqlId(2) = "ADEPTNM"
    UNISqlId(3) = "M6111QA104"
	UNISqlId(4) = "A5144MA103"  '전표입력경로 
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    'rs0에 대한 Value값 setting
    UNIValue(0,0) = lgSelectList
    UNIValue(0,1)  = " " & FilterVar(lgChangeOrgId, "''", "S") & " "
  	UNIValue(0,2)  = " " & FilterVar(lgtxtAccountYear, "''", "S") & " "

	IF lgtxtdeptcd = "" then
		UNIValue(0,3)  = ""
	Else
		UNIValue(0,3)  = " AND A.DEPT_CD =  " & FilterVar(lgtxtdeptcd , "''", "S") & "  "
	end if


	IF lgtxtBizArea = "" then
		UNIValue(0,4)  = strWhere0 & ""
	Else
		UNIValue(0,4)  = strWhere0 & " AND A.BIZ_AREA_CD =  " & FilterVar(lgtxtBizArea, "''", "S") & "  "
	end if

    'rs1에 대한 Value값 setting
	UNIValue(1,0) = " " & FilterVar(lgtxtBizArea, "''", "S") & "  "

	'rs2에 대한 Value값 setting
	IF lgtxtdeptcd = "" then
		UNIValue(2,0)  = "" & FilterVar("XXXXX", "''", "S") & " "				'입력된 값이 없을때 더미값을 넘겨준다
	Else
		UNIValue(2,0)  = " " & FilterVar(lgtxtdeptcd, "''", "S") & " "
	End if
	UNIValue(2,1) = " " & FilterVar(lgChangeOrgId, "''", "S") & " "

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

	'rs3에 대한 Value값 setting
	IF lgtxtCOST_CENTER_CD = "" then
		UNIValue(3,0)  = "''"				                           '입력된 값이 없을때 더미값을 넘겨준다
	Else
		UNIValue(3,0)  = " " & FilterVar(lgtxtCOST_CENTER_CD, "''", "S") & " "
	End if
	UNIValue(4,0) = "" & FilterVar("A1001", "''", "S") & " "
	UNIValue(4,1) = FilterVar(strGlInputType, "''", "S")

	UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'========================================================================================
Sub TrimData()  
	strAmtFr = UNIConvNum(Request("txtAmtFr"),0)
	strAmtTo = UNIConvNum(Request("txtAmtTo"),0)
	lgtxtCOST_CENTER_CD	= Trim(Request("txtCOST_CENTER_CD"))

	strWhere0 = ""

	If strGlInputType <> "" Then
		strWhere0 =  " AND A.GL_INPUT_TYPE =  " & FilterVar(strGlInputType, "''", "S") & " "
	End If

	If StrDesc <> "" Then
		strWhere0 = strWhere0 & " AND A.GL_DESC LIKE " & FilterVar("%" & StrDesc, "''", "S")
	End If
	If strRefNo <> "" Then
		strWhere0 = strWhere0 & " AND A.REF_NO LIKE  " & FilterVar(strRefNo & "%", "''", "S") & ""
	End If

	IF lgtxtCOST_CENTER_CD <> "" then
		strWhere0  = strWhere0 +  " AND A.COST_CD =  " & FilterVar(lgtxtCOST_CENTER_CD, "''", "S") & "  "
	end if
'-------------------------
'금액
'-----------------------
	If strAmtFr <> 0 or strAmtTo <> 0 Then
		If strAmtFr > 0 and strAmtTo <= 0 Then
			strWhere0 = strWhere0 & " and a.DR_LOC_AMT >= " & strAmtFr  & " and a.CR_LOC_AMT >= " & strAmtFr 
		ElseIf strAmtFr <= 0 and strAmtTo > 0 Then
			strWhere0 = strWhere0 & " and a.DR_LOC_AMT <= " & strAmtTo  & " and a.CR_LOC_AMT <= " & strAmtTo 
		Else
			strWhere0 = strWhere0 & " and a.DR_LOC_AMT between " & strAmtFr & " and " & strAmtTo & " and a.CR_LOC_AMT between " & strAmtFr & " and " & strAmtTo
		End If
	End If

End Sub

'========================================================================================
' Query Data
'========================================================================================
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMsgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If  rs0.EOF And rs0.BOF Then
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
 
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
					.Frm1.htxtAccountYear.Value		= "<%=lgtxtAccountYear%>"
					.Frm1.htxtBizArea.Value			= "<%=ConvSPChars(lgtxtBizArea)%>"
					.Frm1.htxtdeptcd.Value			= "<%=ConvSPChars(lgtxtdeptcd)%>"
					.Frm1.horgchangeid.Value		= "<%=ConvSPChars(lgChangeOrgId)%>"
					.Frm1.htxtCOST_CENTER_CD.Value	= "<%=ConvSPChars(lgtxtCOST_CENTER_CD)%>"
					.Frm1.htxtRefNo.Value			= "<%=ConvSPChars(strRefNo)%>"
					.Frm1.htxtDesc.Value			= "<%=ConvSPChars(StrDesc)%>"
					.frm1.htxtGlInputType.value		= "<%=ConvSPChars(strGlInputType)%>"
			End If

        'Show multi spreadsheet data from this line
        .ggoSpread.Source	= .frm1.vspdData      
        .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
        .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag

		.frm1.txtBizAreaNm.value		= "<%=ConvSPChars(biz_area_nm)%>"
		.frm1.txtdeptnm.value			= "<%=ConvSPChars(dept_nm)%>"
		.frm1.txtCOST_CENTER_NM.value	= "<%=ConvSPChars(cost_nm)%>"
		.frm1.txtGlInputType.value		= "<%=ConvSPChars(strGlInputType)%>"
		.frm1.txtGLInputTypeNm.value	= "<%=ConvSPChars(strGlInputTypeNM)%>"
		
		End With
       Parent.DbQueryOk
    Else
	End if
</Script>	

