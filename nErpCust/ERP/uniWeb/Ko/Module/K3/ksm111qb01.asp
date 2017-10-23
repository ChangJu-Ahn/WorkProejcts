<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm111mb01
'*  4. Program Name         : 멀티컴퍼니수주등록 
'*  5. Program Desc         : 멀티컴퍼니수주등록-멀티 
'*  6. Component List       :
'*  7. Modified date(First) : 2005/01/24
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
	Dim lgOpModeCRUD

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
	Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
	Dim iStrPrNo
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount
	Dim lgCurrency
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	Dim lgDataExist
	Dim lgPageNo1
	Dim sRow
	Dim lglngHiddenRows
	Dim MaxRow2
	Dim MaxCount

	Const C_SHEETMAXROWS_D  = 100

	MaxCount = 0
	MaxRow2 = 0
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------

	lgOpModeCRUD  = Request("txtMode")

	'response.write lgOpModeCRUD & lgOpModeCRUD &"<br>"
	'response.write UID_M0001 & UID_M0001 &"<br>"

	'response.end

	Select Case lgOpModeCRUD
	Case CStr(UID_M0001)
	     Call  SubBizQueryMulti()
	End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	lgPageNo1       = UNICInt(Trim(Request("lgPageNo1")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist    = "No"
	iLngMaxRow     = CLng(Request("txtMaxRows"))

	'//Response.write "iLngMaxRow:" & iLngMaxRow


	lgStrPrevKey   = Request("lgStrPrevKey")
	'sRow           = CLng(Request("lRow"))
	'lglngHiddenRows = CLng(Request("lglngHiddenRows"))

	Call FixUNISQLData()
	Call QueryData()

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim strVal
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(0,0)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
	                                                                '    parameter의 수에 따라 변경함 
	UNISqlId(0) = "SM111QA101" 											' header


	UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("strSO_NO"))), " " , "SNM") & "' "

	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
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

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing


	'//Response.write "rs0" & rs0(0) & "<br>"
	'//Response.End

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	Dim FalsechkFlg

	FalsechkFlg = False

	If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		'Call DisplayMsgBox("172400", vbOKOnly, iStrPrNo, "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End
	Else
		Call  MakeSpreadSheetData()
	End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData2 "			& vbCr
	Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr
	Response.Write "	.lgPageNo1		=  """ & lgPageNo1	 & """" & vbCr
	Response.Write "    	.DbQueryOk2(" & MaxCount & ")" & vbCr
	Response.Write "End With"		& vbCr
	Response.Write "</Script>"		& vbCr

End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100
	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt


	Const C_ITEM_CD		= 0	'품목 
	Const C_ITEM_NM		= 1	'품목명 
	Const C_SPEC		= 2	'품목규격 
	Const C_CUST_ITEM_CD	= 3	'고객품목 
	Const C_BP_ITEM_NM	= 4	'고객품목명 
	Const C_BP_ITEM_SPEC	= 5	'고객품목규격 
	Const C_SO_QTY		= 6	'수량 
	Const C_SO_UNIT		= 7	'단위 
	Const C_SO_PRICE	= 8	'단가 
	Const C_NET_AMT2	= 9	'금액 
	Const C_DLVY_DT		= 10	'납기일 
	Const C_VAT_AMT2	= 11	'부가세금액 
	Const C_VAT_RATE2	= 12	'부가세율 
	Const C_VAT_TYPE2	= 13	'부가세유형 
	Const C_VAT_TYPE_NM2	= 14	'부가세유형명 
	Const C_VAT_INC_FLAG	= 15	'부가세포함구분 



	lgDataExist    = "Yes"

	If CLng(lgPageNo1) > 0 Then
		rs0.Move     	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		MaxRow2     	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	End If

	'----- 레코드셋 칼럼 순서 ----------
	'-----------------------------------
	iLoopCount = 0
   	Do while Not (rs0.EOF Or rs0.BOF)
	        iLoopCount =  iLoopCount + 1
	        iRowStr = ""

	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_ITEM_CD))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_ITEM_NM))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SPEC))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_CUST_ITEM_CD))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_ITEM_NM))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_ITEM_SPEC))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_QTY))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_UNIT))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_PRICE))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_NET_AMT2))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_DLVY_DT))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_AMT2))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_RATE2))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_TYPE2))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_TYPE_NM2))
	        'iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_INC_FLAG))

		Select Case ConvSPChars(rs0(C_VAT_INC_FLAG))
		Case "1"
			iRowStr = iRowStr & Chr(11) & "별도"
		Case "2"
			iRowStr = iRowStr & Chr(11) & "포함"
		Case Else
			iRowStr = iRowStr & Chr(11) & ""
		End Select

		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

	        If iLoopCount - 1 < C_SHEETMAXROWS_D Then
	           istrData = istrData & iRowStr & Chr(11) & Chr(12)
	        Else
		   istrData = ""
	           lgPageNo1 = lgPageNo1 + 1

	           Exit Do
	        End If
	        rs0.MoveNext
   	Loop

	'response.write "iLngMaxRow:" & iLngMaxRow & "<br>"
	If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
		lgPageNo1 = ""
	End If

	MaxCount = iLoopCount
	rs0.Close                                                       '☜: Close recordset object
	Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

%>