<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strFrAcqDt
Dim strToAcqDt
Dim strFrAsstNo
Dim strToAsstNo
Dim strAcctCd
Dim strDeptCd

Dim strCond
Dim lgPID

Dim iPrevEndRow
Dim iEndRow
Dim lgDataExist

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
	
	lgPID			= Request("PID") 
	lgPageNo		= Cint(Request("lgStrPrevKey"))                               '☜ : Next key flag
	lgMaxCount		= CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList	= Request("lgSelectList")                               '☜ : select 대상목록 
	lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist		= "No"
	lgTailList		= Request("lgTailList")                                 '☜ : Orderby value

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
    Dim  iLoopCount
    Dim  iRowStr

    lgDataExist    = "Yes"
    lgstrData = ""

    iPrevEndRow = 0

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
		rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
   
    
    Do While Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
 
        If  iLoopCount < lgMaxCount Then
            lgstrData		=	lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop


    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
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
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(0,2)

    If UCase(lgPID) = "A7110MA1" Then									'자산변동내역조회에서는 변동사항이 있는 자산만 보이기 위해.
		UNISqlId(0) = "A7103RA2"
	Elseif UCase(lgPID) = "A7122MA1" Then								'기초자산Master등록 에서 기초자산만 보이도록 
		UNISqlId(0) = "A7103RA3"

	Elseif UCase(lgPID) = "A7103MA1" Then
		UNISqlId(0) = "A7103RA1"

	Elseif UCase(lgPID) = "A7108MA1" Or UCase(lgPID) = "A7109MA1"Then	'고정매각폐기등록화면/고정자산이동등록에서 재고수량이 0인것은 제외 
		UNISqlId(0) = "A7103RA5"
	Else
	    UNISqlId(0) = "A7103RA4"	
	End If

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    'UNIValue(0,2) = UCase(Trim(strToAcqDt)) A7101RA1
    
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
	strFrAcqDt    = UniConvDate(Request("txtFrAcqDt"))
	strToAcqDt    = UniConvDate(Request("txtToAcqDt"))
	strFrAsstNo   = UCase(Trim(Request("txtFrAsstNo")))
	strToAsstNo   = UCase(Trim(Request("txtToAsstNo")))
	strAcctCd     = Request("txtAcctCd") 
	strDeptCd	   = UCase(Trim(Request("txtDeptCd")))
		  
	If strFrAsstNo <> "" Then
	   strCond = strCond & " and A.ASST_NO >=  " & FilterVar( strFrAsstNo, "''", "S") & " "	 
	End If
	     
	If strToAsstNo <> "" Then
	   strCond = strCond & " and A.ASST_NO <=  " & FilterVar(strToAsstNo, "''", "S") & " "
	End If
	         
	If Trim(Request("txtToAcqDt")) <> "" Then
	   strCond = strCond & " and A.REG_DT <=  " & FilterVar(strToAcqDt  , "''", "S") & ""
	End If
	     
	If Trim(Request("txtFrAcqDt")) <> "" Then
	   strCond = strCond & " and A.REG_DT >=  " & FilterVar(strFrAcqDt , "''", "S") & "" 
	End If  
	     
	If strAcctCd <> "" Then
	   strCond = strCond & " and A.ACCT_CD =  " & FilterVar(strAcctCd, "''", "S") & " "
	End If
	     
	If strDeptCd <> "" Then
	   strCond = strCond & " and A.DEPT_CD =  " & FilterVar(strDeptCd, "''", "S") & " "
	End If     
	strCond = strCond & " and A.ORG_CHANGE_ID = C.ORG_CHANGE_ID" 
	     

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	strCond		= strCond	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	


	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data

       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",4),Parent.GetKeyPos("A",6),"A", "Q" ,"X","X")
       Parent.lgStrPrevKey      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If   
</Script>	

