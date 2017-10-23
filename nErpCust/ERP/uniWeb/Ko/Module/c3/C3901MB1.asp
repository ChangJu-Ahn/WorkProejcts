<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call loadInfTB19029B("Q", "C","NOCOOKIE","QB")
Call LoadBasisGlobalInf()

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim lgCodeCond
Dim lgDiffAmt, lgTotAllcAmt

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
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
    
    Const C_SHEETMAXROWS_D = 30
    
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
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(2,3)

    UNISqlId(0) = "C3901MA01"	'sheet1 화면 
    UNISqlId(1) = "C3901MA01"	'Diff Amt
    UNISqlId(2) = "C3901MA03"	'Tot Allc Amt

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    UNIValue(1,0) = "SUM(a.DIFF_AMT)"
    UNIValue(2,0) = "SUM(a.DIFF_AMT)"
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = FilterVar(Trim(Request("txtYyyymm")),"' '","S")
    UNIValue(1,1) = FilterVar(Trim(Request("txtYyyymm")),"' '","S")
    UNIValue(2,1) = FilterVar(Trim(Request("txtYyyymm")),"' '","S")
    
	IF Trim(Request("txtPlantCd")) <> "" Then
		lgCodeCond		= " and a.PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S")
		lgTotAmtCond	= " and a.PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S")
	Else
		lgCodeCond		= ""
		lgTotAmtCond	= ""
	END If

	IF Trim(Request("strProcurType")) <> "" Then
		lgCodeCond		= lgCodeCond   & " and a.TEMP_FG_1 in " & Trim(Request("strProcurType"))
		lgTotAmtCond	= lgTotAmtCond & " and e.PROCUR_TYPE in " & Trim(Request("strProcurType"))
	Else
		lgCodeCond		= lgCodeCond & "" 
		lgTotAmtCond	= lgTotAmtCond & "" 
	END If
    
	IF Trim(Request("txtItemAcct")) <> "" Then
		lgCodeCond		= lgCodeCond   & " and a.TEMP_FG_2 = " & FilterVar(Request("txtItemAcct"), "''", "S")
		lgTotAmtCond	= lgTotAmtCond & " and a.ITEM_ACCT = " & FilterVar(Request("txtItemAcct"), "''", "S")
	Else
		lgCodeCond		= lgCodeCond & "" 
		lgTotAmtCond	= lgTotAmtCond & "" 
	END If
    
	IF Trim(Request("txtItemCd")) <> "" Then
		lgCodeCond		= lgCodeCond   & " and a.ITEM_CD >= " & FilterVar(Request("txtItemCd"), "''", "S")
		lgTotAmtCond	= lgTotAmtCond & " and a.ITEM_CD >= " & FilterVar(Request("txtItemCd"), "''", "S")
	Else
		lgCodeCond		= lgCodeCond & "" 
		lgTotAmtCond	= lgTotAmtCond & "" 
	END If

    UNIValue(0,2) = lgCodeCond
    UNIValue(1,2) = lgCodeCond
    UNIValue(2,2) = lgTotAmtCond
    
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

	lgDiffAmt		= 0
	lgTotAllcAmt	= 0

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
        rs1.Close
        Set rs0 = Nothing
        Set rs1 = Nothing
    Else
		lgDiffAmt = rs1(0)

		If  rs2.EOF And rs2.BOF Then
		    rs2.Close
		    Set rs2 = Nothing
		Else
			lgTotAllcAmt = rs2(0)
		End If

        Call  MakeSpreadSheetData()
    End If
End Sub


%>

<Script Language=vbscript>
	If "<%=lgDataExist%>" = "Yes" Then
			    
		'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
		End If
			    
		Parent.frm1.txtDiffAmt.text		= "<%=UniNumClientFormat(lgDiffAmt,ggAmtOfMoney.Decpoint,0)%>" 
		Parent.frm1.txtTotAllcAmt.text	= "<%=UniNumClientFormat(lgTotAllcAmt,ggAmtOfMoney.Decpoint,0)%>" 

		Parent.ggoSpread.Source  = Parent.frm1.vspdData
		Parent.ggoSpread.SSShowData "<%=lgstrData%>"            '☜ : Display data
		Parent.lgPageNo_A      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		Parent.DbQueryOk("1")
	End If   
</Script>	

