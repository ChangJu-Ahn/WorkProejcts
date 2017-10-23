<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->


<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "C", "NOCOOKIE","MB")

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1							'DBAgent Parameter 선언 


Dim	txtFromYyyymm
Dim	txtToYyyymm
Dim	txtPrevFromYyyymm
Dim	txtPrevToYyyymm

Dim	txtItem
Dim	txtItemNm
Dim txtOptionFlag

Dim lgDataExist
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgErrorStatus
Dim lgStrData
Dim lgKeyStream
Dim strYear
Dim strMonth
Dim strDay
Dim prevStartDate
Dim prevEndDate

									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================

Call HideStatusWnd

On Error Resume Next
Err.Clear

	lgKeyStream      = Split(Request("txtKeyStream"),gColSep)	

    txtFromYyyymm = Replace(Trim(lgKeyStream(0)), gServerDateType ,"")
    txtToYyyymm = Replace(Trim(lgKeyStream(1)), gServerDateType ,"")

    prevStartDate = UNIDateAdd("yyyy",-1,lgKeyStream(0) & gServerDateType & "01",gServerDateFormat)
    Call ExtractDateFrom(prevStartDate,gServerDateFormat,gServerDateType,strYear,strMonth,strDay)
    txtPrevFromYyyymm = strYear & strMonth
    prevEndDate = UNIDateAdd("yyyy",-1,lgKeyStream(1) & gServerDateType & "01",gServerDateFormat)
    Call ExtractDateFrom(prevEndDate,gServerDateFormat,gServerDateType,strYear,strMonth,strDay)
    txtPrevToYyyymm = strYear & strMonth

	
	txtItem	= Trim(lgKeyStream(2))
	txtOptionFlag = Trim(lgKeyStream(3))


    lgDataExist    = "No"
	lgErrorStatus  = "No" 

	Call FixUNISQLData()
	Call QueryData()

	


Sub MakeSpreadSheetData()
    On Error Resume Next
    Dim  iLoopCount

   
    
    lgDataExist    = "Yes"
    lgStrData      = ""

    iLoopCount = 0
    
    IF txtOptionFlag = "Y" Then
    
		Do while Not (rs0.EOF Or rs0.BOF)
		    iLoopCount =  iLoopCount + 1

			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		'손익항목 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(2), ggAmtOfMoney.DecPoint, 0) '금액 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(3), ggExchRate.DecPoint, 0) '매출액대비 비율 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(4), ggAmtOfMoney.DecPoint, 0) 'Prev 금액 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(5), ggExchRate.DecPoint, 0) '매출액대비 비율 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(6), ggAmtOfMoney.DecPoint, 0) '차이금액 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(7), ggExchRate.DecPoint, 0) '비율 
				
			lgstrData = lgstrData & Chr(11) & iLoopCount 
			lgstrData = lgstrData & Chr(11) & Chr(12)		
				
		    
		    rs0.MoveNext
		Loop
	ELSE '단가 
		Do while Not (rs0.EOF Or rs0.BOF)
		    iLoopCount =  iLoopCount + 1

			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		'손익항목 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(2), ggQty.DecPoint, 0) '금액 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(3), ggExchRate.DecPoint, 0) '매출액대비 비율 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(4), ggQty.DecPoint, 0) 'Prev 금액 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(5), ggExchRate.DecPoint, 0) '매출액대비 비율 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(6), ggQty.DecPoint, 0) '차이금액 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(7), ggExchRate.DecPoint, 0) '비율 
				
			lgstrData = lgstrData & Chr(11) & iLoopCount 
			lgstrData = lgstrData & Chr(11) & Chr(12)		
				
		    
		    rs0.MoveNext
		Loop
	END IF
		
	
    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = 0													'☜: 다음 데이타 없다.
    End If

  	
	rs0.Close
    Set rs0 = Nothing 
    
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------


    Redim UNIValue(1,23)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

	IF txtOptionFlag = "Y" Then
		UNISqlId(0) = "GE005MA01"
		
		UNIValue(0,0) = FilterVar(txtFromYYYYMM, "''", "S")
		UNIValue(0,1) = FilterVar(txtToYYYYMM, "''", "S")
		UNIValue(0,2) = FilterVar(txtItem, "''", "S")
		UNIValue(0,3) = FilterVar(txtFromYYYYMM, "''", "S")
		UNIValue(0,4) = FilterVar(txtToYYYYMM, "''", "S")
		UNIValue(0,5) = FilterVar(txtItem, "''", "S")
		UNIValue(0,6) = FilterVar(txtFromYYYYMM, "''", "S")
		UNIValue(0,7) = FilterVar(txtToYYYYMM, "''", "S")
		UNIValue(0,8) = FilterVar(txtItem, "''", "S")
		UNIValue(0,9) = FilterVar(txtPrevFromYYYYMM, "''", "S")
		UNIValue(0,10) = FilterVar(txtPrevToYYYYMM, "''", "S")
		UNIValue(0,11) = FilterVar(txtItem, "''", "S")
		UNIValue(0,12) = FilterVar(txtPrevFromYYYYMM, "''", "S")
		UNIValue(0,13) = FilterVar(txtPrevToYYYYMM, "''", "S")
		UNIValue(0,14) = FilterVar(txtItem, "''", "S")
		UNIValue(0,15) = FilterVar(txtPrevFromYYYYMM, "''", "S")
		UNIValue(0,16) = FilterVar(txtPrevToYYYYMM, "''", "S")
		UNIValue(0,17) = FilterVar(txtItem, "''", "S")
	ELSE
		UNISqlId(0) = "GE005MA02"	
	
		UNIValue(0,0) = FilterVar(txtFromYYYYMM, "''", "S")
		UNIValue(0,1) = FilterVar(txtToYYYYMM, "''", "S")
		UNIValue(0,2) = FilterVar(txtItem, "''", "S")
		UNIValue(0,3) = FilterVar(txtFromYYYYMM, "''", "S")
		UNIValue(0,4) = FilterVar(txtToYYYYMM, "''", "S")
		UNIValue(0,5) = FilterVar(txtItem, "''", "S")
		UNIValue(0,6) = FilterVar(txtFromYYYYMM, "''", "S")
		UNIValue(0,7) = FilterVar(txtToYYYYMM, "''", "S")
		UNIValue(0,8) = FilterVar(txtItem, "''", "S")
		UNIValue(0,9) = FilterVar(txtFromYYYYMM, "''", "S")
		UNIValue(0,10) = FilterVar(txtToYYYYMM, "''", "S")
		UNIValue(0,11) = FilterVar(txtItem, "''", "S")
		UNIValue(0,12) = FilterVar(txtPrevFromYYYYMM, "''", "S")
		UNIValue(0,13) = FilterVar(txtPrevToYYYYMM, "''", "S")
		UNIValue(0,14) = FilterVar(txtItem, "''", "S")
		UNIValue(0,15) = FilterVar(txtPrevFromYYYYMM, "''", "S")
		UNIValue(0,16) = FilterVar(txtPrevToYYYYMM, "''", "S")
		UNIValue(0,17) = FilterVar(txtItem, "''", "S")
		UNIValue(0,18) = FilterVar(txtPrevFromYYYYMM, "''", "S")
		UNIValue(0,19) = FilterVar(txtPrevToYYYYMM, "''", "S")
		UNIValue(0,20) = FilterVar(txtItem, "''", "S")
		UNIValue(0,21) = FilterVar(txtPrevFromYYYYMM, "''", "S")
		UNIValue(0,22) = FilterVar(txtPrevToYYYYMM, "''", "S")
		UNIValue(0,23) = FilterVar(txtItem, "''", "S")
	END IF
	
        	
	UNISqlId(1) = "COMMONQRY"					'ITEM_NM
    UNIValue(1,0)  = "SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD=" & FilterVar(txtItem, "''", "S") 
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    


    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
        
        
    IF NOT (rs1.EOF or rs1.BOF) then
		txtItemNm = rs1(0)				
	ELSE
		txtItemNm = ""
	End if
    rs1.Close
    Set rs1 = Nothing 

    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else
        Call  MakeSpreadSheetData()
    End If


    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
		
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
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub


%>

<Script Language=vbscript>
 
	With Parent
	   
	   .frm1.txtDeptNm.value	= "<%=ConvSPChars(txtItemNm)%>"
	   
		If "<%=lgDataExist%>" = "Yes" AND "<%=lgErrorStatus%>" <> "YES" Then
		   .ggoSpread.Source  = Parent.frm1.vspdData
		   .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
		   .DbQueryOk
		End If

    End With

</Script>	
	

<%
Set lgADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
