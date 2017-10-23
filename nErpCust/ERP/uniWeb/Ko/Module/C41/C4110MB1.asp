<% Option Explicit 
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : C4110MB1
'*  4. Program Name         : 원가차이반영 
'*  5. Program Desc         : 기존 재고차이, 입고차이를 통합 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2007/03/01
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Cho Ig Sung
'* 10. Modifier (Last)      :
'* 11. Comment              :
'=======================================================================================================-->
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->

<%
	Server.ScriptTimeOut = 30000
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")     
    
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

    Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD
'	Dim lgLngMaxRow,   lgMaxCount

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data

Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim lgCodeCond

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet

'	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'	lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData



    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQueryMulti()

        Case "ExeReflect"
             Call ExeReflect()

        Case "ExeCancel"
             Call ExeCancel()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList	= Request("lgSelectList")                               '☜ : select 대상목록 
	lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgTailList		= Request("lgTailList")                                 '☜ : Orderby value
	lgDataExist		= "No"

	Call FixUNISQLData()
	Call QueryData()

End Sub


Sub FixUNISQLData()

	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 

	Redim UNIValue(2,3)

	UNISqlId(0) = "C4110MA1"		'sheet1 화면 

	UNIValue(0,0) = lgSelectList                                          '☜: Select list

	lgCodeCond		= " and a.yyyymm = " & FilterVar(Trim(Request("txtYyyymm")),"' '","S")
	    
	IF Trim(Request("txtPlantCd")) <> ""	Then	lgCodeCond		= lgCodeCond   & " and a.PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S")

	IF Trim(Request("txtCostCd")) <> ""		Then	lgCodeCond		= lgCodeCond   & " and a.cost_Cd = " & FilterVar(Request("txtCostCd"), "''", "S")

	IF Trim(Request("txtTrnsTypeCd")) <> ""	Then	lgCodeCond		= lgCodeCond   & " and a.trns_Type = " & FilterVar(Request("txtTrnsTypeCd"), "''", "S")

	IF Trim(Request("txtMovTypeCd")) <> ""	Then	lgCodeCond		= lgCodeCond   & " and a.mov_type = " & FilterVar(Request("txtMovTypeCd"), "''", "S")

	IF Trim(Request("txtItemAcctCd")) <> ""	Then	lgCodeCond		= lgCodeCond   & " and a.item_acct = " & FilterVar(Request("txtItemAcctCd"), "''", "S")

	IF Trim(Request("txtItemCd")) <> ""		Then	lgCodeCond		= lgCodeCond   & " and a.ITEM_CD = " & FilterVar(Request("txtItemCd"), "''", "S")

	IF Trim(Request("txtTrackingNo")) <> ""	Then	lgCodeCond		= lgCodeCond   & " and a.Tracking_No = " & FilterVar(Request("txtTrackingNo"), "'*'", "S")

	UNIValue(0,1) = lgCodeCond

	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub


Sub QueryData()
	Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
	Dim iStr
	Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
	    
	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If
	    
	If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		rs0.Close
		Set rs0 = Nothing
	Else
		Call  MakeSpreadSheetData()
	End If

End Sub

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 5000
    
    lgstrData = ""

    lgDataExist    = "Yes"

    If CInt(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CInt(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
       
    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
				
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub ExeReflect()
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
				    
	Dim iPC4G060
	Dim	I1_yyyymm 

	I1_yyyymm = Trim(Request("txtYyyymm"))

	Set iPC4G060 = Server.CreateObject("PC4G060.cCMngCostDiffSvr")

	If CheckSYSTEMError(Err, True) = True Then					
		Call SetErrorStatus
		Exit Sub
	End If  
				    
	Call iPC4G060.C_REFLECT_COST_DIFF_SVR(gStrGloBalCollection, I1_yyyymm)

	If CheckSYSTEMError(Err, True) = True Then					
		Call SetErrorStatus
		Set iPC4G060 = Nothing
		Exit Sub
	End If    
				    
	Set iPC4G060 = Nothing
End Sub    


Sub ExeCancel()
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
				    
	Dim iPC4G060
	Dim	I1_yyyymm 

	I1_yyyymm = Trim(Request("txtYyyymm"))

	Set iPC4G060 = Server.CreateObject("PC4G060.cCMngCostDiffSvr")

	If CheckSYSTEMError(Err, True) = True Then					
		Call SetErrorStatus
		Exit Sub
	End If  
				    
	Call iPC4G060.C_CANCEL_COST_DIFF_SVR(gStrGloBalCollection, I1_yyyymm)

	If CheckSYSTEMError(Err, True) = True Then					
		Call SetErrorStatus
		Set iPC4G060 = Nothing
		Exit Sub
	End If    
				    
	Set iPC4G060 = Nothing

End Sub




'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

<Script Language="VBScript">
    
	Select Case "<%=lgOpModeCRUD %>"
		Case "<%=UID_M0001%>"                                                         '☜ : Query
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				If "<%=lgDataExist%>" = "Yes" Then
					Parent.ggoSpread.Source		= Parent.frm1.vspdData
					Parent.ggoSpread.SSShowData	"<%=lgstrData%>"			'☜ : Display data
					Parent.lgPageNo				= "<%=lgPageNo%>"			'☜ : Next next data tag

					Parent.DBQueryOk        
				End If   
			End If
		Case "<%="ExeReflect"%>"
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				Parent.ExeReflectOk
			End If   
		Case "<%="ExeCancel"%>"
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				Parent.ExeCancelOk
			End If      
	End Select    
    
       
</Script>	
