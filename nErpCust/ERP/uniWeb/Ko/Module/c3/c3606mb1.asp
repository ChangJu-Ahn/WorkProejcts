<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")       

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4          '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim txtCostCd
Dim txtCostNm
Dim txtYYYYMM
Dim txtTotAmt															'⊙ : 배부대상합계 
Dim txtTotWorkinAmtSum													'⊙ : 총 재공배부합계액 
Dim txtTotItemAmtSum													'⊙ : 총 배부금액 

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = Trim(Request("lgPageNo"))                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
'   lgMaxCount     = Trim(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    
	txtCostCd = Trim(Request("txtCostCd"))
	txtYYYYMM = Trim(Request("txtYYYYMM"))

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

    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim strWhere
	Dim strWhere1
	Dim strWhere2

    Redim UNIValue(4,3)

    UNISqlId(0) = "C3606MA101"
    UNISQLID(1) = "commonqry"
    UNISqlId(2) = "C3606MA103"
    UNISqlId(3) = "C3606MA106"
    UNISqlId(4) = "C3606MA107"


    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    
    strWhere = ""
    If txtCostCd <> "" Then
		strWhere = strWhere & " AND b.cost_cd = " & FilterVar(txtCostCd ,"''"	,"S")		'코스트센타코드 
	End If
	strWhere = strWhere & " AND c.yyyymm = " & FilterVar(txtYYYYMM ,"''"	,"S")		'년월 
	
	strWhere2 = strWhere
	'strWhere = strWhere & " group by b.cost_cd,b.cost_nm,a.acct_cd,a.acct_nm,isnull(d.minor_nm,'*') "
	
	strWhere1 = ""
    strWhere1 = " yyyymm = " & FilterVar(txtYYYYMM ,"''"	,"S") 
    
    IF txtCostCd <> "" Then
		strWhere1 = strWhere1 & " AND cost_cd = " & FilterVar(txtCostCd ,"''"	,"S") 
    END IF

	UNIValue(0,1) = FilterVar(txtYYYYMM ,"''"	,"S")
    UNIValue(0,2) = strWhere	          '공장코드 
    UNIValue(1,0) = "select Cost_Nm from B_COST_CENTER where cost_cd = " & FilterVar(txtCostCd ,"''"	,"S")
    
    UNIValue(2,0) = FilterVar(txtYYYYMM ,"''"	,"S")
    UNIValue(2,1) = strWhere2
    
    UNIValue(3,0) = strWhere1
    UNIValue(4,0) = strWhere1

    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	on Error Resume Next
	Err.Clear 
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
	iStr = Split(lgstrRetMsg,gColSep)
	
 
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    'rs1
	If txtCostCd <> "" Then
	    If Not (rs1.EOF OR rs1.BOF) Then
			txtCostNm = Trim(rs1("Cost_Nm"))
		Else
			txtCostNm = ""
			Call DisplayMsgBox("124400", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs1.Close
		    Set rs1 = Nothing 
			Exit sub
		End IF
		rs1.Close
	    Set rs1 = Nothing
	End If

     If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("233500", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()

	    'rs2
		If Not (rs2.EOF OR rs2.BOF) Then
			txtTotAmt = Trim(rs2(0))
		Else
			txtTotAmt = ""
		End IF
		rs2.Close
		Set rs2 = Nothing

		'rs3'⊙ : 총 재공배부 합계액 
		If Not (rs3.EOF OR rs3.BOF) Then
			txtTotWorkinAmtSum = Trim(rs3(0))
		Else
			txtTotWorkinAmtSum = ""
		End IF
		rs3.Close
		Set rs3 = Nothing
			
			
		'rs4'⊙ :  총 배부합계액 
		If Not (rs4.EOF OR rs4.BOF) Then
			txtTotItemAmtSum = Trim(rs4(0))
		Else
			txtTotItemAmtSum = ""
		End IF
		rs4.Close
		Set rs4 = Nothing
    End If
 
    
End Sub

%>
<Script Language=vbscript>
With Parent

	.frm1.txtTotAmt.text = ""
	.frm1.txtTotWorkinAmtSum.text = "<%=UNINumClientFormat(0,ggAmtOfMoney.Decpoint,0)%>"			
	.frm1.txtTotItemAmtSum.text = "<%=UNINumClientFormat(0,ggAmtOfMoney.Decpoint,0)%>"			

    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          .Frm1.htxtCostCd.Value = .Frm1.txtCostCd.Value                  'For Next Search
          .Frm1.htxtYYYYMM.Value = "<%=txtYYYYMM%>"                  'For Next Search
       End If


		.frm1.txtTotAmt.text = "<%=UNINumClientFormat(txtTotAmt,ggAmtOfMoney.Decpoint,0)%>"			'rs2 값(배부대상금액합계)
		.frm1.txtTotWorkinAmtSum.text = "<%=UNINumClientFormat(txtTotWorkinAmtSum,ggAmtOfMoney.Decpoint,0)%>"			'rs2 값(배부대상금액합계)
		.frm1.txtTotItemAmtSum.text = "<%=UNINumClientFormat(txtTotItemAmtSum,ggAmtOfMoney.Decpoint,0)%>"			'rs2 값(배부대상금액합계)

       .ggoSpread.Source  = .frm1.vspdData
       .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
       .lgPageNo_A      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       .DbQueryOk("1")
    End If

	.frm1.txtCostNm.value = "<%=ConvSPChars(txtCostNm)%>"			'rs1 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 

End With
</Script>
