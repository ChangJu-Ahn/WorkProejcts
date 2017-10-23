<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : Asset Acquisition Reference Popup
'*  3. Program ID           : A7107rb1.asp
'*  4. Program Name         : 자산변동 참조 팝업 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/06/10
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Kim Hee Jung
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Expires = -1                                                       '☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True                                                     '☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgPID                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strFrChgDt	                                                           
Dim strToChgDt
Dim strFrChgNo	                                                           
Dim strToChgNo
Dim strDeptCd
Dim strFrAsstNo	                                                           
Dim strToAsstNo
	       
Dim strCond

' 권한관리 추가 
Dim lgAuthBizAreaCd	' 사업장 
Dim lgInternalCd	' 내부부서 
Dim lgSubInternalCd	' 내부부서(하위포함)
Dim lgAuthUsrID		' 개인 
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

    lgPID          = UCase(Request("PID"))  
    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
             iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(0,2)
	
	if lgPID = "A7109MA1" THEN
		UNISqlId(0) = "A7107RA2"
	ELSE
		UNISqlId(0) = "A7107RA1"
	END IF
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    'UNIValue(0,2) = UCase(Trim(strToChgDt)) A7101RA1
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		rs0.Close
		Set rs0 = Nothing
		Set lgADF = Nothing
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     strFrChgNo	  = Request("txtFrChgNo")                                                          
     strToChgNo   = Request("txtToChgNo")     
     strFrChgDt   = UniConvDate(Request("txtFrChgDt"))
     strToChgDt   = UniConvDate(Request("txtToChgDt"))
     strDeptCd	  = Request("txtDeptCd")     
     strFrAsstNo  = Request("txtFrAsstNo")                                                          
     strToAsstNo  = Request("txtToAsstNo")     
     
     If strFrChgNo <> "" Then
		strCond = strCond & " and A.CHG_NO >=  " & FilterVar(strFrChgNo , "''", "S") & ""	 
     End If
     
     If strToChgNo <> "" Then
		strCond = strCond & " and A.CHG_NO <=  " & FilterVar(strToChgNo , "''", "S") & ""
     End If
     	 
     If Trim(Request("txtToChgDt")) <> "" Then
		strCond = strCond & " and A.CHG_DT <=  " & FilterVar(strToChgDt  , "''", "S") & ""
     End If
     
     If Trim(Request("txtFrChgDt")) <> "" Then
		strCond = strCond & " and A.CHG_DT >=  " & FilterVar(strFrChgDt  , "''", "S") & ""
     End If  
     
     If strDeptCd <> "" Then
		strCond = strCond & " and A.FROM_DEPT_CD =  " & FilterVar(strDeptCd , "''", "S") & ""
     End If
     
     If strFrAsstNo <> "" Then
		strCond = strCond & " and A.ASST_CD >=  " & FilterVar(strFrAsstNo , "''", "S") & ""	 
     End If
     
     If strToAsstNo <> "" Then
		strCond = strCond & " and A.ASST_CD <=  " & FilterVar(strToAsstNo , "''", "S") & ""
     End If          
	
	 IF lgPID = "A7107MA1" Then	
		strCond = strCond & " And A.CHG_FG IN (" & FilterVar("01", "''", "S") & " ," & FilterVar("02", "''", "S") & " )"
	 elseIF lgPID = "A7108MA1" Then	
		strCond = strCond & " And a.CHG_FG in (" & FilterVar("03", "''", "S") & " ," & FilterVar("04", "''", "S") & " )"	
	 END IF		 
	 
	 strCond = strCond & " and isnull(a.asst_chg_no," & FilterVar("","''","S") & ") = " & FilterVar("","''","S") 

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND c.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND c.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND c.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND c.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If   
		 
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub


%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
         .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '☜: set next data tag
         .DbQueryOk
	End with
</Script>	

