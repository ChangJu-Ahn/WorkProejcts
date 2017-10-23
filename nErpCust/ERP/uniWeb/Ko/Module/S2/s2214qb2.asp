<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2214QB2
'*  4. Program Name         : 판매계획대실적조회(품목그룹)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/16
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Yong Sik
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                         '☜ : DBAgent Parameter 선언 
    Dim lgstrData                                                              '☜ : data for spreadsheet data
    Dim lgStrPrevKey                                                           '☜ : 이전 값 
    Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgDataExist
    Dim lgPageNo
    Dim lgStrColorFlag, lgStrDisplayType, lgStrGrpNm
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgPageNo		= ""
    lgSelectList   = Request("txtSelectLIst")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("txtSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("txtTailList")                                 '☜ : Orderby value
	lgStrDisplayType = Request("cboDisplayType")

    lgDataExist    = "No"

    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 100     
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    iLoopCount = 0
    lgStrColorFlag = ""
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
		If rs0(0) > 0 Then
			lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & rs0(0) & gRowSep
		End If
		
		lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        rs0.MoveNext
	Loop

	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim iStrFromDt, iStrToDt, iStrLocExpFlag, iStrItemGroupCd, iStrBaseCur, iStrCur
	Dim iIntGrpLvl
	
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,15)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	iStrFromDt		= UNIGetFirstDay(Request("txtConFromDt"),gDateFormatYYYYMM)
	iStrToDt		= UNIGetLastDay(Request("txtConToDt"), gDateFormatYYYYMM)
	iStrLocExpFlag	= Request("cboConLocExpFlag")
	iIntGrpLvl		= Request("cboConGrpLvl")
	iStrItemGroupCd = Trim(Request("txtConItemGroupCd"))
	iStrBaseCur		= Request("cboConBaseCur")
	iStrCur			= Trim(Request("txtConCur"))
	
	If lgStrDisplayType = "H" Then
	    UNISqlId(0) = "S2214QA201"
	    lgSelectList = Replace(lgSelectList, "1?", "총계")
	    lgSelectList = Replace(lgSelectList, "2?", "그룹소계")
	Else
	    UNISqlId(0) = "S2214QA202"
		lgSelectList = Replace(lgSelectList, "1?", "총계")
		lgSelectList = Replace(lgSelectList, "2?", "년소계")
		lgSelectList = Replace(lgSelectList, "3?", "월소계")
		lgSelectList = Replace(lgSelectList, "4?", "그룹소계")
	End If

	UNIValue(0,0) = lgSelectList

	UNIValue(0,1) = " " & FilterVar(UNIConvDate(iStrFromDt), "''", "S") & ""			' 시작일 
	UNIValue(0,2) = " " & FilterVar(UNIConvDate(iStrToDt), "''", "S") & ""			' 종료일 
	UNIValue(0,3) = "" & FilterVar("IG", "''", "S") & ""										' 품목그룹별 조회 
	UNIValue(0,4) = "NULL"										' 영업조직레벨 
	UNIValue(0,5) = "NULL"										' 품목그룹 
	UNIValue(0,6) = "NULL"										' 사용여부 
	
	UNIValue(0,7)  = iIntGrpLvl									' 품목레벨 
	If iStrItemGroupCd = "" Then								' 품목그룹 
		UNIValue(0,8) = "NULL"
	Else
		UNIValue(0,8) = " " & FilterVar(iStrItemGroupCd, "''", "S") & ""
		UNISqlId(1) = "I224QA1A5"								' 품목그룹명 Fetch
		UNIValue(1,0) = UNIValue(0,8) & " AND ITEM_GROUP_LEVEL = " & iIntGrpLvl
	End If
	
	UNIValue(0,9) = "" & FilterVar("%", "''", "S") & ""										' 사용여부 

	UNIValue(0,10) = " " & FilterVar(Request("cboConSpType"), "''", "S") & ""		' 판매계획유형 

	If iStrLocExpFlag = "" Then									' 내수/수출여부 
		UNIValue(0,11) = "" & FilterVar("%", "''", "S") & ""
	Else
		UNIValue(0,11) = " " & FilterVar(iStrLocExpFlag, "''", "S") & ""
	End If

	UNIValue(0,12) = " " & FilterVar(iStrBaseCur, "''", "S") & ""					' 화폐기준 
	
	If iStrCur <> "" Then										' 화폐단위 
		UNIValue(0,13) = " " & FilterVar(iStrCur, "''", "S") & ""
		If iStrBaseCur = "D" Then
			If lgStrDisplayType = "H" Then
				UNIValue(0,14) = "AND GROUPING_FLAG <> 1 "		' 그룹별 소계는 제외시킴 
			Else
				UNIValue(0,14) = "WHERE GROUPING_FLAG <> 1 "
			End If
		Else
			UNIValue(0,14) = ""
		End If
		
		' 화폐존재여부 Check
		UNISqlId(2) = "s0000qa014"
		UNIValue(2,0) = FilterVar(iStrCur, "''", "S")
	Else
		UNIValue(0,13) = "" & FilterVar("%", "''", "S") & ""
		UNIValue(0,14) = ""
	End If
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If

	' Write the Script Tag(<Script language=vbscript>)
	Call BeginScriptTag()

    If  UNIValue(1,0) <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing
			Call ConNotFound("txtConItemGroupCd")
			Exit Sub
		Else
			Call WriteConDesc("txtConItemGroupNm", Rs1(0))
		End If
	Else
		Call WriteConDesc("txtConItemGroupNm", "")
    End If
    
    ' 화폐 존재여부    
    If  UNIValue(2,0) <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing
			Call ConNotFound("txtConCur")
			Exit Sub
		End If
    End If

    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("cboConSpType")
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
        Call WriteResult()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Write the Result
' 결과Html을 작성한다.
'----------------------------------------------------------------------------------------------------------
Sub BeginScriptTag()
	Response.Write "<Script language=VBScript> " & VbCr
End Sub

Sub EndScriptTag()
	Response.Write "</Script> " & VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성(조회조건 포함)
Sub ConNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1." & pvStrField & ".alt, ""X"") " & VbCr
	Response.Write "Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회조건에 해당하는 명을 Display하는 Script 작성 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write "Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성 
Sub DataNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write "Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	If lgStrDisplayType = "H" Then
		Response.Write "Parent.ggoSpread.Source  = Parent.frm1.vspdData " & VbCr
		Response.Write  "Parent.frm1.vspdData.Redraw = False  "      & vbCr      	
		Response.Write  "Parent.ggoSpread.SSShowDataByClip   """ & lgstrData & """ ,""F""" & vbCr
		Response.Write "parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
		Response.Write "Parent.DbQueryOk " & VbCr	
		Response.Write "Parent.frm1.vspdData.Redraw = True  "       & vbCr      
	Else
		Response.Write "Parent.ggoSpread.Source  = Parent.frm1.vspdData2 " & VbCr
		Response.Write  "Parent.frm1.vspdData2.Redraw = False  "      & vbCr      	
		Response.Write  "Parent.ggoSpread.SSShowDataByClip   """ & lgstrData & """ ,""F""" & vbCr
		Response.Write "parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
		Response.Write "Parent.DbQueryOk " & VbCr
		Response.Write "Parent.frm1.vspdData2.Redraw = True  "       & vbCr      
	End If	
	Call EndScriptTag()
End Sub
%>
