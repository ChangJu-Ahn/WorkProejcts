<%Option Explicit%>
<%
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3601mb1
'*  4. Program Name         : CC별 배부내역 조회 
'*  5. Program Desc         : CC별 배부내역 조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/27
'*  8. Modified date(Last)  : 2002/03/25
'*  9. Modifier (First)     : Cho Ig Sung
'* 10. Modifier (Last)      : JANG YOON KI
'* 11. Comment              :
'=======================================================================================================

%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")       

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                             '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
DIM txtYyyyMm
Dim txtCostCd
DIM txtCostNm
Dim txtOriginTotAmt
Dim SetFocusFlag
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = Trim(Request("lgPageNo"))                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
'   lgMaxCount     = Trim(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
	
	txtYyyyMm = Trim(Request("txtYyyyMm"))
	txtCostCd = Trim(Request("txtCostCd"))
	
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
    
   	Const C_SHEETMAXROWS_D  = 100 
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time

    lgstrData = ""

    lgDataExist    = "Yes"

    If UniConvNumStringToDouble(lgPageNo,0) > 0 Then
       rs0.Move     = UniConvNumStringToDouble(lgMaxCount,0) * UniConvNumStringToDouble(lgPageNo,0)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
       
    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
				
        If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = Cstr(UniConvNumStringToDouble(lgPageNo,0) + 1)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim strWhere
	
	
	
    Redim UNIValue(2,2)

    UNISqlId(0) = "C3601MA101"
    UNISqlId(1) = "C3601MA104"
    UNISqlId(2) = "CommonQry"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
    
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    
    strWhere = " and YYYYMM = " & FilterVar(txtYyyyMm , "''", "S")
	
    if txtCostCd <> "" then
		strWhere = strWhere & " and GIVE_COST_CD = " & FilterVar(txtCostCd   , "''", "S")
	end if
	
	UNIValue(0,1) = strWhere
    UNIValue(1,0) = strWhere    
    UNIValue(2,0) = "select cost_nm from b_cost_center Where cost_cd= " & FilterVar(txtCostCd, "''","S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
     iStr = Split(lgstrRetMsg,gColSep)
    
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If    
    
    If  rs0.EOF AND rs0.BOF Then
		SetFocusFlag = 2
		Call DisplayMsgBox("233500", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		rs0.close
		set rs0 = Nothing
    Else 
		SetFocusFlag = 0   
        Call  MakeSpreadSheetData()
		
    End If    

	If Not (rs1.EOF OR rs1.BOF) Then
		txtOriginTotAmt = Trim(rs1(0))
	Else
		txtOriginTotAmt = "0"
	End IF		

	If Not (rs2.EOF OR rs2.BOF) Then
		txtCostNm = Trim(rs2("Cost_Nm"))
	Else
		SetFocusFlag = 1
		txtCostNm = ""
	End IF		
	
   	rs1.Close
	rs2.Close

    Set rs1 = Nothing
    Set rs2 = Nothing
	
End Sub


%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then    
		With Parent
		   'Set condition data to hidden area
		   If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
		      .Frm1.txtCostCd.Value = .Frm1.txtCostCd.Value                  'For Next Search
		   End If
										
		   .ggoSpread.Source  = .frm1.vspdData
		   .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
		   .lgPageNo_A      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		   .DbQueryOk("1")
		   .Frm1.txtOriginTotAmt.text = "<%=UNINumClientFormat(txtOriginTotAmt, ggAmtOfMoney.DecPoint, 0)%>"
		
		End With
	Else		
		If <%=SetFocusFlag%> = 1 Then 
			parent.Frm1.txtCostCd.Focus		
		Else
			parent.Frm1.txtYyyyMm.Focus		
		End if
		Parent.Frm1.txtOriginTotAmt.text = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint, 0)%>"
		Parent.Frm1.txtAmtSum.text = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint, 0)%>"
		Parent.Frm1.txtAllocAmtSum.text = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint, 0)%>"
    End If  
    Parent.Frm1.txtCostNm.Value = "<%=ConvSPChars(txtCostNm)%>"
    
    
</Script>
