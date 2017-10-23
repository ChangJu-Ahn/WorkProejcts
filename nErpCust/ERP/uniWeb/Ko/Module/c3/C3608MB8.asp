<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf() 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim lgPlantNm
Dim lgFrItemNm
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

	
    lgPageNo       = Trim(Request("lgPageNo"))                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
'   lgMaxCount     = Trim(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
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

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

  	Call CommonQueryRs("PLANT_NM","B_PLANT","PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)	
			'공장이 존재하지 않습니다.
		Exit Sub
	Else
		lgPlantNm = Trim(Replace(lgF0,Chr(11),""))
	End if

  	Call CommonQueryRs("ITEM_NM","B_ITEM","ITEM_CD = " & FilterVar(Request("txtFrItemCd"), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		lgFrItemNm = ""
	Else
		lgFrItemNm = Trim(Replace(lgF0,Chr(11),""))
	End if


    Redim UNIValue(0,4)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "c3608ma101"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	UNIValue(0,1)  = FilterVar(Request("txtPlantCd"),"''"       ,"S")               '학교 
	UNIValue(0,2)  = FilterVar(Request("txtYyyymm")   ,"" & FilterVar("zzzzzzz", "''", "S") & "","S")               '학년 
	UNIValue(0,3)  = FilterVar(Request("txtFrItemCd"),"''"       ,"S")               '학교 
    
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
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    
    

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If    
        
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
	Parent.Frm1.txtPlantNM.Value	  = "<%=ConvSPChars(lgPlantNm)%>"
	Parent.Frm1.txtFrItemNM.Value	  = "<%=ConvSPChars(lgFrItemNm)%>"

    If "<%=lgDataExist%>" = "Yes" Then
		
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.hPlantCd.Value      = "<%=Request("txtPlantCd")%>"                  'For Next Search
          Parent.Frm1.hFrItemCd.Value      = "<%=Request("txtFrItemCd")%>"                  'For Next Search
		  
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
    End If   

</Script>	

