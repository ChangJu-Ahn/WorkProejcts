<%@ LANGUAGE=VBSCript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%

On Error Resume Next
Err.Clear

Call HideStatusWnd																	'☜: Hide Processing message
Dim lgOpModeCRUD

Dim ADF																				'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg																		'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0,rs1,rs2,rs3,rs4,rs5,rs6,rs7,rs8		'DBAgent Parameter 선언 
Dim i,strCond,strMsgCd,strMsg1
Dim strDt
Dim strCostCd
Dim strItemCd
Dim strFctrCd
Dim lgstrData																		'☜ : data for spreadsheet data
Dim lgPageNo
Dim iPrevEndRow
Dim iEndRow
Dim lgTailList																		'☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist

Const C_SHEETMAXROWS_D = 1000

	Call HideStatusWnd																'☜: Hide Processing message
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE","MB")  	

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)							'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    lgSelectList   = Request("lgSelectList")										'☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)						'☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")											'☜ : Orderby value
    lgDataExist    = "No"
    iPrevEndRow    = 0
    iEndRow        = 0    

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()

Sub MakeSpreadSheetData()
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr

    lgstrData = ""

    lgDataExist    = "Yes"

    If CInt(lgPageNo) > 0 Then
		iPrevEndRow = C_SHEETMAXROWS_D * CDbl(lgPageNo)    
		rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
       
    iLoopCount = -1

    Do While Not (rs0.EOF Or rs0.BOF)
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
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If

	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub  FixUNISQLData()

    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3, 2)
	

	
    UNISqlId(0) = "C3820MA1"
	UNISqlId(1) = "CommonQry"
	UNISqlId(2) = "CommonQry"
	UNISqlId(3) = "CommonQry"
 
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    UNIValue(0,1) = strCond
   
	UNIValue(1,0) = " select COST_CD,COST_NM from B_COST_CENTER where COST_CD = " & FilterVar(UCase(strCostCd),"''","S")
    UNIValue(2,0) = " select DSTB_FCTR_CD,DSTB_FCTR_NM from C_DSTB_FCTR where DSTB_FCTR_CD = " & FilterVar(UCase(strFctrCd),"''","S")
    UNIValue(3,0) = " select ITEM_CD,ITEM_NM from B_ITEM where ITEM_CD = " & FilterVar(UCase(strItemCd),"''","S")
    
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
  
End Sub

Sub TrimData()
'Dim strDt
'Dim strCostCd
'Dim strFctrCd,strItemCd
	strDt = Trim(REPLACE(Request("txtDt"),"-",""))
	strCostCd = Trim(Request("txtCostCd"))
	strFctrCd = Trim(Request("txtFctrCd"))
	strItemCd = Trim(Request("txtItemCd"))
	strCond = ""

	strCond = strCond & " and A.YYYYMM = '" & strDt & "'"	
	
	If strCostCd <> "" Then 
		strCond = strCond & " and C.COST_CD = " & FilterVar(strCostCd,"''","S")
	End If
		
	If strFctrCd <> "" Then 
		strCond = strCond & " and B.DSTB_FCTR_CD = " & FilterVar(strFctrCd,"''","S")
	End If
	
	If strItemCd <> "" Then 
		strCond = strCond & " and A.PRNT_ITEM_CD = " & FilterVar(strItemCd,"''","S")
	End If
		strCond = strCond & "  GROUP BY A.DSTB_FCTR_CD,B.DSTB_FCTR_NM,A.COST_CD,C.COST_NM,A.PRNT_PLANT_CD,D.PLANT_NM,E.MINOR_NM,A.PRNT_ITEM_CD,F.ITEM_NM   "
End Sub

Sub  QueryData()
    Dim iStr

	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4,rs5,rs6,rs7,rs8)
	Set ADF = Nothing

    iStr = Split(strRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(strRetMsg , vbInformation, I_MKSCRIPT)
		Response.End 	
    End If    

	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strCostCd <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtCostCd_Alt")
		End If
	Else		
%>
	<Script Language=vbscript>
		parent.frm1.txtCostCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
		parent.frm1.txtCostNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
	</Script>
<%
	End If
			
	rs1.Close
	Set rs1 = Nothing

	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And Trim(Request("txtFctrCd")) <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtFctrCd_Alt")
		End If
	Else	
%>
	<Script Language=vbscript>
		parent.frm1.txtFctrCd.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
		parent.frm1.txtFctrNm.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
	</Script>
<%		
	End If
			
	rs2.Close
	Set rs2 = Nothing
			
	If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And Trim(Request("txtItemCd")) <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtItemCd_Alt")
		End If
	Else		
%>
	<Script Language=vbscript>
		parent.frm1.txtItemCd.value = "<%=ConvSPChars(Trim(rs3(0)))%>"
		parent.frm1.txtItemNm.value = "<%=ConvSPChars(Trim(rs3(1)))%>"
	</Script>
<%		
	End If		

	rs3.Close
	Set rs3 = Nothing

		If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1,"" , I_MKSCRIPT)
		Response.End	
	End If	
	
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else  
        Call  MakeSpreadSheetData()
    End If
    
End Sub    

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
    
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists

       End If
    
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk()
    End If  
</Script>	