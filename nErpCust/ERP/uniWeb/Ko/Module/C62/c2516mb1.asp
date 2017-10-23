<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->

<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 표준원가관리 
'*  3. Program ID           : c2516mb1
'*  4. Program Name         : 표준원가 재료비 산출근거조회 
'*  5. Program Desc         : 공장별 표준계산시 재료비에 대한 산출근거를 한다.
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2002/03/22
'*  8. Modified date(Last)  : 2002/03/
'*  9. Modifier (First)     : Eun Hee, Hwang
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Call LoadBasisGlobalInf() 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                             '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
DIM ITEM_NM
DIM PLANT_NM
 
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 
    lgPageNo       = Trim(Request("lgPageNo"))                 '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
'    lgMaxCount     = Trim(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
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
    
    Const C_SHEETMAXROWS_D  = 100 
	
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time

    lgDataExist    = "Yes"
    lgstrData      = ""
	
    If UniConvNumStringToDouble(lgPageNo,0) > 0 Then
       rs0.Move     = UniConvNumStringToDouble(lgMaxCount,0) * UniConvNumStringToDouble(lgPageNo,0)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
     
    'rs0에 대한 결과 
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
	DIm strWhere
	
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(2,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "C2510MA01"
    UNISqlId(1) = "COMMONQRY"					'PLANT_NM
    UNISqlId(2) = "COMMONQRY"					'ITEM_NM

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strWhere = " AND A.PLANT_CD = "
	strWhere = strWhere & " " & FilterVar(Request("txtPlantCd"), "''", "S") & " "
	
	If Request("txtCItemCd") <>"" then
		strWhere= strWhere & " AND A.ITEM_CD >="
		strWhere = strWhere & " " & FilterVar(Trim(Request("txtCItemCd"))   , "''", "S") & " " 
	End If
	
	UNIValue(0,1)  = strWhere 
	
	UNIValue(1,0)  = "SELECT PLANT_NM FROM B_PLANT WHERE PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S") & " "		'rs1에 대한 Value값 setting
	UNIValue(2,0)  = "SELECT A.Item_nm from b_item A, b_item_by_plant B WHERE A.item_cd = B.item_cd AND A.item_cd =  " & FilterVar(Request("txtCItemCd"), "''", "S") & " "		'rs2에 대한 Value값 setting
	      
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF
                                                                 '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
	
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
	iStr = Split(lgstrRetMsg,gColSep)
    
    'If iStr(0) <> "0" Then
     '   Call ServerMesgBox(iStr(1), vbInformation, I_MKSCRIPT)
    'End If    
	 
	 'rs1에 대한 결과 
    IF NOT (rs1.EOF or rs1.BOF) then
		PLANT_NM = rs1("PLANT_NM")
	ELSE
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
%>
	<Script Language=vbscript>
		parent.frm1.txtPlantCd.focus
	</Script>
<%
        rs1.Close
        Set rs1 = Nothing
        Exit Sub
	End if
    rs1.Close
    Set rs1 = Nothing 
    
    'rs2에 대한 결과 
    IF NOT (rs2.EOF or rs2.BOF) then
	    ITEM_NM = rs2("ITEM_NM")
	ELSE
		ITEM_NM=""
    END IF
    rs2.Close
    Set rs2 = Nothing

    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("232100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
 %>
	<Script Language=vbscript>
		parent.frm1.txtCItemCD.focus
	</Script>
 <%
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
		
        Call  MakeSpreadSheetData()
    End If
  
End Sub

%>

<Script Language=vbscript>
	With Parent
		If "<%=lgDataExist%>" = "Yes" Then
		   'Set condition data to hidden area
		   If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.hPlantCd.value = "<%=Request("txtPlantCd")%>"
				.frm1.hCItemCd.value = "<%=Request("txtCItemCd")%>"
		   End If
		   'Show multi spreadsheet data from this line
		   .ggoSpread.Source  = Parent.frm1.vspdData
		   .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
		   .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		   .DbQueryOk
		End If   
		 .frm1.txtPlantNM.value = "<%=ConvSPChars(PLANT_NM)%>"
		 .frm1.txtCItemNM.value = "<%=ConvSPChars(ITEM_NM)%>"
	End With
</Script>	


