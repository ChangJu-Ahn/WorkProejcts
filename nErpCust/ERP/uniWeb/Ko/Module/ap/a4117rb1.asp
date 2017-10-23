
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

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3                         '☜ : DBAgent Parameter 선언 
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
	Dim strFrApDt	                                                           
	Dim strToApDt
	Dim strFrApNo	                                                           
	Dim strToApNo
	Dim strdeptcd
	DIm strDealBpCd
	Dim stradjustNo
	Dim strCond

	Dim strMsgCd
	Dim strMsg1

' 권한관리 추가 
Dim lgAuthBizAreaCd	' 사업장 
Dim lgInternalCd	' 내부부서 
Dim lgSubInternalCd	' 내부부서(하위포함)
Dim lgAuthUsrID		' 개인	
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    


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

	Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(3,2)

    UNISqlId(0) = "A4101RA101"
    UNISqlId(1) = "ADEPTNM"
	UNISqlId(2) = "ABPNM"
	UNISqlId(3) = "Commonqry"
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(" " & FilterVar(strDeptCd, "''", "S") & " " )
    UNIValue(1,1) = UCase(" " & FilterVar(UCase(Request("txtOrgChangeId")), "''", "S") & " " )
  
    UNIValue(2,0) = FilterVar(UCase(strDealBpCd), "''", "S") 
    UNIValue(3,0) = "SELECT adjust_no FROM A_AP_ADJUST WHERE adjust_no=" & FilterVar(UCase(stradjustNo), "''", "S") 
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strDeptCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDeptCd_Alt")
%>
<Script Language=vbScript>
		With parent
			.frm1.txtDeptNm.value = ""
		End With
		</Script>
<%
		Else
%>
<Script Language=vbScript>
		With parent
			.frm1.txtDeptNm.value = ""
		End With
		</Script>
<%
		End If
    Else
%>
<Script Language=vbScript>
		With parent
			.frm1.txtDeptCd.value = "<%=Trim(ConvSPChars(rs1(0)))%>"
			.frm1.txtDeptNm.value = "<%=Trim(ConvSPChars(rs1(1)))%>"
		End With
		</Script>
<%
    End If
    
	Set rs1 = Nothing 
 
 
 If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strDealBpCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDealBpCd_Alt")
		End If
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtDealBpNm.value = ""
			End With
		</Script>
<%		
	Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtDealBpNm.value = "<%=Trim(ConvSPChars(rs2(1)))%>"
			End With
		</Script>
<%		
	End IF
	rs2.Close
	Set rs2 = Nothing

    
    If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And stradjustNo <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtadjustNo_Alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtadjustNo.focus
		End With
		</Script>
<%
	Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtadjustNo.value = "<%=Trim(ConvSPChars(rs3(0)))%>"
		End With
		</Script>
<%
    End If
    rs3.Close
	Set rs3 = Nothing 
	
	
    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    End If
            
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     strFrApDt     = UCase(Trim(UNIConvDate(Request("txtFrApDt"))))
     strToApDt     = UCase(Trim(UNIConvDate(Request("txtToApDt"))))
     strFrApNo	   = UCase(Trim(Request("txtFrApNo")))
     strToApNo     = UCase(Trim(Request("txtToApNo")))
     strdeptcd     = UCase(Trim(Request("txtdeptcd")))
     strDealBpCd   = UCase(Trim(Request("txtDealBpCd")))
     stradjustNo	= UCase(Trim(Request("txtadjustNo")))
     
     
     
     strCond  = " and A.conf_fg = " & FilterVar("C", "''", "S") & "  AND A.GL_NO <>''  AND (A.AP_STS = " & FilterVar("O", "''", "S") & "  OR A.AP_NO IN (SELECT DISTINCT AP_NO FROM A_AP_ADJUST)) "
     
	If strFrApDt <> "" Then
	   strCond = strCond & " and A.AP_DT >=  " & FilterVar(strFrApDt , "''", "S") & ""
	End If
	     
	If strToApDt <> "" Then
	   strCond = strCond & " and A.AP_DT <=  " & FilterVar(strToApDt , "''", "S") & ""
	End If
	     
	If strFrApNo <> "" Then
	   strCond = strCond & " and A.AP_NO >=  " & FilterVar(strFrApNo , "''", "S") & ""
	End If
	     
	If strToApNo <> "" Then
	   strCond = strCond & " and A.AP_NO <=  " & FilterVar(strToApNo , "''", "S") & ""
	End If
	     
	If strdeptcd <> "" Then
	   strCond = strCond & " and A.dept_cd =  " & FilterVar(strdeptcd , "''", "S") & ""
	End If     
	     
	If strDealBpCd <> "" Then	
	   strCond = strCond & " And A.deal_bp_cd = " & FilterVar(strDealBpCd , "''", "S") 
	End If
		 
	If stradjustNo <> "" Then	
	   strCond = strCond & " AND A.AP_NO = (SELECT distinct AP_NO FROM A_AP_ADJUST WHERE adjust_no=" & FilterVar(stradjustNo , "''", "S") & ")" 
	End if

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
          Parent.Frm1.htxtFrApDt.Value      = Parent.Frm1.txtFrApDt.Text
          Parent.Frm1.htxtToApDt.Value         = Parent.Frm1.txtToApDt.Text
          Parent.Frm1.htxtFrApNo.Value  = Parent.Frm1.txtFrApNo.Value
          Parent.Frm1.htxtToApNo.Value    = Parent.Frm1.txtToApNo.Value
          Parent.Frm1.htxtdeptcd.Value    = Parent.Frm1.txtdeptcd.Value
          Parent.Frm1.htxtDealBpCd.Value	= Parent.Frm1.txtDealBpCd.Value
          Parent.Frm1.htxtadjustNo.Value	= Parent.Frm1.txtadjustNo.Value
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",6),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",14),"A", "Q" ,"X","X")
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
    End If   

</Script>	

