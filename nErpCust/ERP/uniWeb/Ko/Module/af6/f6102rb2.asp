
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
Call LoadInfTB19029B("*", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                        '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim iPrevEndRow
Dim iEndRow	


'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
	Dim strPPDt 
    Dim strToPPDt
    Dim strDocCur
    Dim strBpCd
    Dim strDeptCd

	Dim strCond

	Dim strMsgCd
	Dim strMsg1

' 권한관리 추가 
Dim lgAuthBizAreaCd	' 사업장 
Dim lgInternalCd	' 내부부서 
Dim lgSubInternalCd	' 내부부서(하위포함)
Dim lgAuthUsrID		' 개인	
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
	Dim BP_NM
	Dim DEPT_NM
  
    
    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0

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
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
       rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
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

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(2,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "F6102RA201"
	UNISqlId(1) = "ABPNM"
    UNISqlId(2) = "ADEPTNM"
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,1)  = strCond
    UNIValue(1,0) = FilterVar(strBpCd, "''", "S")
    
    UNIValue(2,0) = FilterVar(strDeptCd, "''", "S")
    UNIValue(2,1) = FilterVar(UCase(Request("txtOrgChangeId")), "''", "S")    

      
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	Dim iStr
	Dim strMsg
    Dim strMsg1
    Dim strMsgCd
    Dim strMsgCd1
    
    strMsg = Trim(Request("txtBpcd_Alt"))
    strMsg1 = Trim(Request("txtDeptCd_Alt"))
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	IF NOT (rs1.EOF or rs1.BOF) then
		BP_NM = rs1(0)
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = "<%=BP_NM%>"   
		End With
		</Script>
<%			
	ELSE
		if Trim(Request("txtBpCd")) <> "" Then
			strMsgCd = "970000"
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = ""   
		End With
		</Script>
<%	
		Else 
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = ""   
		End With
		</Script>
<%			
		End if
	End if
    rs1.Close
    Set rs1 = Nothing 
    
    'rs2에 대한 결과 
    IF NOT (rs2.EOF or rs2.BOF) then
	    DEPT_NM = rs2(0)
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtDeptNm.Value = "<%=DEPT_NM%>"   
		End With
		</Script>
<%			    
	ELSE
		if Trim(Request("txtDeptCd")) <> "" Then
			strMsgCd1 = "970000"
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtDeptNm.Value = ""   
		End With
		</Script>
<%	
		Else
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtDeptNm.Value = ""   
		End With
		</Script>
<%		
		End if
    END IF
    rs2.Close
    Set rs2 = Nothing
	
	If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg, "", I_MKSCRIPT)
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    End If
    
    If  "" & Trim(strMsgCd1) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    End If
	
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Response.End										
	Else
		Call  MakeSpreadSheetData()
    End If				
    
    Set rs0 = Nothing
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     strPPDt		= UCase(Trim(UNIConvDate(Request("txtPPDt"))))
     strToPPDt     = UCase(Trim(UNIConvDate(Request("txtToPPDt"))))
     strDocCur	   = UCase(Trim(Request("txtDocCur")))                                                          
     strBpCd		= UCase(Trim(Request("txtBpCd")))
     strDeptCd     = UCase(Trim(Request("txtDeptCd")))
   
     If strPPDt <> "" Then strCond = strCond & " and A.PRPAYM_DT >= " & FilterVar(strPPDt, "''", "S") 
    
     If strToPPDt <> "" Then strCond = strCond & " and A.PRPAYM_DT <= " & FilterVar(strToPPDt, "''", "S")

     If strDocCur <> "" Then strCond = strCond & " and A.DOC_CUR = " & FilterVar(strDocCur, "''", "S")

     If strDeptCd <> "" Then strCond = strCond & " AND A.DEPT_CD = " & FilterVar(strDeptCd, "''", "S")
   
     If strBpCd <> "" Then strCond = strCond & " AND A.BP_CD = " & FilterVar(strBpCd, "''", "S")

	 strCond = strCond & " AND A.PRPAYM_STS  = " & FilterVar("O", "''", "S") & "   "
	 strCond = strCond & " AND A.CONF_FG = " & FilterVar("C", "''", "S") & "  "
	 strCond = strCond & " AND A.bal_amt <> 0 "
	 strCond = strCond & " AND A.GL_NO <> " & FilterVar("","''","S") & "  "	  

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If     
		  
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub
%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.htxtDeptCd.Value      = Parent.Frm1.txtDeptCd.Value
          Parent.Frm1.htxtBpCd.Value         = Parent.Frm1.txtBpCd.Value
          Parent.Frm1.htxtDocCur.Value  = Parent.Frm1.txtDocCur.Value
          Parent.Frm1.htxtPPDt.Value    = Parent.Frm1.txtPPDt.Text
          Parent.Frm1.htxtToPPDt.Value    = Parent.Frm1.txtToPPDt.Text
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",8),Parent.GetKeyPos("A",12),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",8),Parent.GetKeyPos("A",14),"A", "Q" ,"X","X")
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
    End If   
</Script>	

