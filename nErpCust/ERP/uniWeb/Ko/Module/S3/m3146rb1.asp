<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2000/11/01
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Const C_SHEETMAXROWS_D  = 30                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Dim strPoType	                                                           '⊙ : 발주형태 
Dim strPoFrDt	                                                           '⊙ : 발주일 
Dim strPoToDt	                                                           '⊙ :
Dim strSpplCd	                                                           '⊙ : 공급처 
Dim strPurGrpCd	                                                           '⊙ : 구매그룹 
Dim strItemCd	                                                           '⊙ : 품목 
Dim strTrackNo	                                                           '⊙ : Tracking No
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)    
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
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
	
    If  iLoopCount < C_SHEETMAXROWS_D Then                             '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing                                       '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(0,2)

    UNISqlId(0) = "m3146ra101"									'* : 데이터 조회를 위한 SQL문 
 
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " AND A.TRACKING_NO = " & FilterVar(Request("txtTrackingNo"), "''", "S") & " "		

    UNIValue(0,1) = strVal   '---발주일 
   
'================================================================================================================   
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0) '* : Record Set 의 갯수 조정 
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
	If  rs0.EOF And rs0.BOF Then
	    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	    rs0.Close
	    Set rs0 = Nothing
		Response.End     
	Else    
	    Call  MakeSpreadSheetData()
	End If	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub
%>
<Script Language=vbscript>
    With parent
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '☜: Display data 
        .lgPageNo = "<%=lgPageNo%>"
        .DbQueryOk
	End with
</Script>	

