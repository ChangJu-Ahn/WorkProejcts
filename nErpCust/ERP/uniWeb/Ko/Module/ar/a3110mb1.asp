<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Err.Clear 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1,rs2, rs3, rs4, rs5, rs6                '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iPrevEndRow
Dim iEndRow	


Dim strBpCd
Dim strBpNM
Dim strDeptCd
Dim strDeptNm
Dim strFromDt
Dim strToDt
Dim strCostCd
Dim strCOST_CENTER_NM
Dim strcboArSts
Dim strcboConfFg
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1
Dim strProject
Dim strCond

Dim txtTotArLocAmt
Dim txtTotClsLocAmt
Dim txtTotBalLocAmt
Dim iChangeOrgId


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL



'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
	Call HideStatusWnd()
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A","NOCOOKIE","MB")
	Call LoadBNumericFormatB("Q", "A","NOCOOKIE","MB")         

	lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList	= Request("lgSelectList")                               '☜ : select 대상목록 
	lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgTailList		= Request("lgTailList")                                 '☜ : Orderby value
	lgDataExist		= "No"
	iPrevEndRow		= 0
	iEndRow			= 0
	    
	iChangeOrgId = Trim(request("OrgChangeId"))
	    
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

	Const C_SHEETMAXROWS_D = 100
    
    If CInt(lgPageNo) > 0 Then
		iPrevEndRow = C_SHEETMAXROWS_D * CDbl(lgPageNo)    
       rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If

  	
	rs0.Close
    Set rs0 = Nothing     

    If Not( rs1.EOF OR rs1.BOF) Then		
		txtTotArLocAmt = rs1(0)
		txtTotClsLocAmt = rs1(1)
		txtTotBalLocAmt = rs1(2)
	Else
		txtTotArLocAmt = 0
		txtClsAmt = 0
		txtBalAmt = 0
    End IF
    
    rs1.Close
    Set rs1 = Nothing 
    
     'rs2
    If Not( rs2.EOF OR rs2.BOF) Then
		strDeptCd = Trim(rs2(0))
		strDeptNm = Trim(rs2(1))
	Else
		strDeptCd = ""
		strDeptNm = ""
    End IF
    
    rs2.Close
    Set rs2 = Nothing 
    
     'rs3
    If Not( rs3.EOF OR rs3.BOF) Then
		strBpCd = Trim(rs3(0))
		strBpNm = Trim(rs3(1))
	Else
		strBpCd = ""
		strBpNm = ""
    End IF
    
    rs3.Close
    Set rs3 = Nothing 
    
     ' rs4
    If Not( rs4.EOF OR rs4.BOF) Then
   		strCostCd = Trim(rs4(0))
		strCOST_CENTER_NM = Trim(rs4(1))
	Else
		strCostCd = ""
		strCOST_CENTER_NM = ""
		
    End IF
    
    rs4.Close
    Set rs4 = Nothing 
    
    ' rs5
    If Not( rs5.EOF OR rs5.BOF) Then
   		strBizAreaCd = Trim(rs5(0))
		strBizAreaNm = Trim(rs5(1))
	Else
		strBizAreaCd = ""
		strBizAreaNm = ""
		
    End IF
    
    rs5.Close
    Set rs5 = Nothing
    
    ' rs6
    If Not( rs6.EOF OR rs6.BOF) Then
   		strBizAreaCd1 = Trim(rs6(0))
		strBizAreaNm1 = Trim(rs6(1))
	Else
		strBizAreaCd1 = ""
		strBizAreaNm1 = ""
		
    End IF
    
    rs6.Close
    Set rs6 = Nothing
    
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim TempSql
    Redim UNISqlId(6)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(6,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "A3110MA101"
    UNISqlId(1) = "A3110MA102"
    UNISqlId(2) = "ADEPTNM"
    UNISqlId(3) = "ABPNM"
	UNISqlId(4) = "commonqry"
	UNISqlId(5) = "A_GETBIZ"
    UNISqlId(6) = "A_GETBIZ"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,1) = UCase(Trim(strCond))
	UNIValue(1,0) = UCase(Trim(strCond))
	
	UNIValue(2,0)  = " " & FilterVar(strDeptCd, "''", "S") & " "		
	UNIValue(2,1)  = " " & FilterVar(iChangeOrgId, "''", "S") & " "	
	UNIValue(3,0)  = " " & FilterVar(strBpCd, "''", "S") & " "
	UNIValue(4,0)  = " select cost_cd, cost_nm from B_COST_CENTER where cost_cd =  " & FilterVar(strCostCd, "''", "S") & " "
	UNIValue(5,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(6,0)  = FilterVar(strBizAreaCd1, "''", "S")
	    
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
        
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3, rs4, rs5, rs6)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Where Query Data
'---------------------------------------------------------------------------------------------------------
Sub TrimData()
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strBpCd				= Trim(Request("txtBpCd"))
	strDeptCd			= Trim(Request("txtDeptCd"))
	strFromDt			= UNIConvDate(Trim(Request("txtFromDt")))
	strToDt				= UNIConvDate(Trim(Request("txtToDt")))
	strCostCd			= Trim(Request("txtCostCd"))
	strcboArSts			= Trim(Request("cboArSts"))
	strcboConfFg		= Trim(Request("cboConfFg"))
	strBizAreaCd		= Trim(UCase(Request("txtBizAreaCd")))            '사업장From
	strBizAreaCd1		= Trim(UCase(Request("txtBizAreaCd1")))            '사업장To
	strProject			= Trim(Request("txtProject"))
	 
	strCond = " AND A.ar_dt >=  " & FilterVar(strFromDt , "''", "S") & " AND A.ar_dt <=  " & FilterVar(strToDt , "''", "S") & "" 
	
	strCond = strCond & " AND A.conf_fg = " & FilterVar(strcboConfFg ,null,"S")
	
	If strBpCd <> ""		Then strCond = strCond & " AND A.deal_bp_cd = "	& FilterVar(strBpCd ,null,"S")
	
	If strDeptCd <> ""		Then 
		strCond = strCond & " AND A.ORG_CHANGE_ID =  " & FilterVar(iChangeOrgId , "''", "S") & " AND A.DEPT_CD = " & FilterVar(strDeptCd ,null,"S")
	End if
	
	If strCostCd <> ""		Then strCond = strCond & " AND A.cost_cd =  " & FilterVar(strCostCd ,null,"S")
	
	If strcboArSts <> ""	Then strCond = strCond & " AND A.ar_sts = "		& FilterVar(strcboArSts ,null,"S")
	
	
	
	if strBizAreaCd <> ""	then
		strCond = strCond & " AND A.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strCond = strCond & " AND A.BIZ_AREA_CD >= "	& FilterVar("", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> ""	then
		strCond = strCond & " AND A.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strCond = strCond & " AND A.BIZ_AREA_CD <= "	& FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if
	
	If strProject <> ""		Then	strCond = strCond & " AND A.project_no Like  " & FilterVar("%" & strProject & "%", "''", "S") & "" 


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
       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			With parent.frm1

				.hBpCd.value = Trim(.txtBpCd.value)
				.hDeptCd.value =  Trim(.txtDeptCd.value)
				.hFromDt.value =  Trim(.txtFromDt.text)
				.hToDt.value = Trim(.txtToDt.text)
				.hConfFg.value = Trim(.cboArSts.value)
				.hArSts.value = Trim(.cboConfFg.value)
				.htxtCostCd.value= Trim(.txtCostCd.value)
			End With
       End If
       
       'Show multi spreadsheet data from this line
       With Parent
		.ggoSpread.Source  = Parent.frm1.vspdData
		Parent.frm1.vspdData.Redraw = False
		Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",4),Parent.GetKeyPos("A",1),"A", "Q" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",4),Parent.GetKeyPos("A",2),"A", "Q" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",4),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",4),Parent.GetKeyPos("A",5),"D", "Q" ,"X","X")		
		Parent.frm1.vspdData.Redraw = True	
		.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		
		.frm1.txtDeptCd.value = "<%=strDeptCd%>"
		.frm1.txtDeptNm.value = "<%=strDeptNm%>"
		.frm1.txtBpCd.value = "<%=strBpCd%>"
		.frm1.txtBpNm.value = "<%=strBpNm%>"
		.frm1.txtCostCd.value="<%=strCostCd%>"
		.frm1.txtCostNm.value="<%=strCOST_CENTER_NM%>"
		.frm1.txtBizAreaCd.value="<%=strBizAreaCd%>"
		.frm1.txtBizAreaNm.value="<%=strBizAreaNm%>"
		.frm1.txtBizAreaCd1.value="<%=strBizAreaCd1%>"
		.frm1.txtBizAreaNm1.value="<%=strBizAreaNm1%>"
		.frm1.txtTotArLocAmt =  "<%=UNINumClientFormat(txtTotArLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
        .frm1.txtTotClsLocAmt = "<%=UNINumClientFormat(txtTotClsLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
        .frm1.txtTotBalLocAmt = "<%=UNINumClientFormat(txtTotBalLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.htxtBizAreaCd.value="<%=strBizAreaCd%>"
		.frm1.htxtBizAreaCd1.value="<%=strBizAreaCd1%>"		        
        .DbQueryOk
       End With
    End If   

</Script>	

