<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>


<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Open Acct Connection
'*  3. Program ID           : a5403ra2
'*  4. Program Name         : 미결연결팝업 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2002/10/23
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : Jeong Yong Kyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Expires = -1                                                       '☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True                                                     '☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")
Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4			'☜ : DBAgent Parameter 선언 
Dim lgStrPrevKey															'☜ : 이전 값 
Dim lgTailList																'☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim lgStrflag
Dim lgtxtAcctCd
Dim lgtxtFromDt
Dim lgtxtToDt
Dim lgtxtDocCur
Dim lgtxtBizCd		'>>air
Dim lgtxtNo
Dim lgtxtMgntCd1
Dim lgtxtMgntCd2
Dim lgtxtGlNoSeq
Dim lgtxtMaxRows
Dim lgtxtGlNo

Dim strFrDueDt	'air 
Dim strToDueDt	'air

Dim strMsgCd
Dim strMsg1
Dim skip_rs4,skip_rs5,no_mgnt1,no_mgnt2

Const C_SHEETMAXROWS_D  = 30 
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)										'☜: Max fetched data at a time
    lgPageNo       = Request("lgPageNo")									'☜ : Next key flag
    lgSelectList   = Request("lgSelectList")								'☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)				'☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")									'☜ : Orderby value
    lgDataExist    = "No"
 
    Call SubOpenDB(lgObjConn)                                               '☜: Make a DB Connection
    Call TrimData()  
    Call FixUNISQLData()
    Call QueryData()
    Call SubCloseDB(lgObjConn)                                              '☜: Close DB Connection    
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
	lgtxtAcctCd		= Trim(Request("txtAcctCd"))
	lgtxtFromDt		= Request("txtFromDt")
	lgtxtToDt		= Request("txtToDt")	
	
	lgtxtDocCur		= Trim(Request("txtDocCur"))
	lgtxtBizCd		= Trim(Request("txtBizCd"))
'Call ServerMesgBox(lgtxtBizCd , vbInformation, I_MKSCRIPT)	
	lgtxtGlNo		= Trim(Request("txtGlNo"))
	lgtxtMgntCd1	= Trim(Request("txtMgntCd1"))
	lgtxtMgntCd2	= Trim(Request("txtMgntCd2"))
	lgtxtGlNoSeq	= Trim(Request("txtGlNoSeq"))
	lgStrflag		= Request("Strflag")
	lgtxtMaxRows	= Request("txtMaxRows")
	
	strFrDueDt		= Trim(Request("txtFrDueDt"))
	strToDueDt  	= Trim(Request("txtToDueDt"))
	
	if strToDueDt="" then strToDueDt="2999-12-31"	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim ii
	Dim iArrTemp
	Dim iArrTemp2
	Dim iStrWhere0
	Dim iStrWhere
	Dim iIntCnt
	Dim IntRetCD   
	Dim temp_nm1,temp_nm2
	Dim stbl_id,scol_id,stbl_id2,scol_id2,sMajor_cd,sMajor_cd2

	Redim UNISqlId(5)                                                    '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(5,10)                                                      '⊙: DB-Agent로 전송될 parameter를 위한 변수 

	UNISqlId(0) = "a5403ra201"
	UNISqlId(1) = "CommonQry"
	UNISqlId(2) = "CommonQry"
	UNISqlId(3) = "CommonQry"
	UNISqlId(4) = "CommonQry"
	
	UNIValue(0,0) = lgSelectList  
  	
  	If lgtxtFromDt <>  "" Then
	   iStrWhere0 = " a.gl_dt >=  " & FilterVar(lgtxtFromDt , "''", "S") & " AND "  
	End If 
		
	If lgtxtToDt <>  "" Then
	   iStrWhere0 = iStrWhere0 & " a.gl_dt <=  " & FilterVar(lgtxtToDt , "''", "S") & " AND "  
	End If 
			
	If lgtxtAcctCd <>  "" Then
	   iStrWhere0 = iStrWhere0 & "  a.acct_cd =  " & FilterVar(lgtxtAcctCd , "''", "S") & " AND "  
	End If 
	
	If lgtxtDocCur <>  "" Then
	   iStrWhere0 = iStrWhere0 & "  a.doc_cur =  " & FilterVar(lgtxtDocCur , "''", "S") & " AND "  
	End If
'Call ServerMesgBox(lgtxtBizCd , vbInformation, I_MKSCRIPT)	
	'>>air	
	If lgtxtBizCd <> "" Then
		'Call ServerMesgBox("A" , vbInformation, I_MKSCRIPT)	
		iStrWhere0 = iStrWhere0 & "	b.biz_area_cd = " & FilterVar(lgtxtBizCd , "''", "S") & " AND "
	End If	
	 
    	
	If lgtxtGlNo <> "" Then
		iStrWhere0 = iStrWhere0 & "   a.gl_no LIKE   " & FilterVar(lgtxtGlNo & "%", "''", "S") & " AND "  
	End If
	
	If lgtxtMgntCd1 <> "" Then
		iStrWhere0 = iStrWhere0 & "   a.mgnt_val1 LIKE   " & FilterVar(lgtxtMgntCd1 & "%", "''", "S") & " AND "  
	End If
	
	If lgtxtMgntCd2 <> "" Then
		iStrWhere0 = iStrWhere0 & "   a.mgnt_val2 LIKE   " & FilterVar(lgtxtMgntCd2 & "%", "''", "S") & " AND "  
	End If

	'AIR
  	If strFrDueDt <>  "" Then
	   iStrWhere0 = iStrWhere0 & " a.due_dt >=  " & FilterVar(strFrDueDt , "''", "S") & " AND "  
	End If 
		
	If strToDueDt <>  "" Then
	   iStrWhere0 = iStrWhere0 & " a.due_dt <=  " & FilterVar(strToDueDt , "''", "S") & " AND "  
	End If
	'AIR
'Call ServerMesgBox(iStrWhere0 , vbInformation, I_MKSCRIPT)		
	UNIValue(0,1) = iStrWhere0			

	UNIValue(1,0) = " select a.acct_cd,a.acct_nm from a_acct a inner join a_acct_gp b on a.gp_cd=b.gp_cd "
	UNIValue(1,0) = UNIValue(1,0) & " where a.del_fg <> " & FilterVar("Y", "''", "S") & "  and a.mgnt_fg = " & FilterVar("Y", "''", "S") & "  and a.mgnt_type = " & FilterVar("9", "''", "S") & "  "
	UNIValue(1,0) = UNIValue(1,0) & "  and  a.acct_cd =  " & FilterVar(lgtxtAcctCd , "''", "S") & ""
	UNIValue(2,0) = " select gl_no from a_gl where gl_no =  " & FilterVar(lgtxtGlNo , "''", "S") & ""
	
	If lgtxtMgntCd1 <> "" Then
		lgStrSQL = ""
		lgStrSQL = lgStrSQL & " select b.tbl_id,b.DATA_COLM_ID,b.data_colm_nm,ISNULL(LTRIM(RTRIM(b.MAJOR_CD)),'')MAJOR_CD from a_acct a , a_ctrl_item b "
		lgStrSQL = lgStrSQL & " where a.mgnt_cd1 = b.ctrl_cd and a.acct_cd =  " & FilterVar(lgtxtAcctCd , "''", "S") & ""
		        
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			no_mgnt1 = "TRUE"
			UNIValue(3,0) = " "
		Else	
			If lgStrflag = 1 then
				stbl_id = Trim(lgObjRs("tbl_id"))
				scol_id = Trim(lgObjRs("DATA_COLM_ID"))
				temp_nm1 = Trim(lgObjRs("data_colm_nm"))
				sMajor_cd = Trim(lgObjRs("MAJOR_CD"))

				UNIValue(3,0) = " select distinct(a.mgnt_val1),b."&temp_nm1 & " from a_open_acct a , "&stbl_id & " b"
				UNIValue(3,0) = UNIValue(3,0) & " where a.acct_cd =  " & FilterVar(lgtxtAcctCd , "''", "S") & ""
				UNIValue(3,0) = UNIValue(3,0) & "  and  a.mgnt_val1 = b."&scol_id		
				UNIValue(3,0) = UNIValue(3,0) & "  and  a.mgnt_val1 =  " & FilterVar(lgtxtMgntCd1 , "''", "S") & ""
				If sMajor_cd <> "" then
					UNIValue(3,0) = UNIValue(3,0) & "  and  b.major_cd =  " & FilterVar(sMajor_cd , "''", "S") & ""
                                End If  
			Else
				skip_rs4 = "TRUE"
				UNIValue(3,0) = " "
			End if
		End If

		Call SubCloseRs(lgObjRs)															'☜: Release RecordSSet	
	Else		
		skip_rs4 = "TRUE"
		UNIValue(3,0) = " "
	End If

	If lgtxtMgntCd2 <> "" Then
		lgStrSQL = ""
		lgStrSQL = lgStrSQL & " select b.tbl_id,b.DATA_COLM_ID,b.data_colm_nm,ISNULL(LTRIM(RTRIM(b.MAJOR_CD)),'')MAJOR_CD from a_acct a , a_ctrl_item b "
		lgStrSQL = lgStrSQL & " where a.mgnt_cd2 = b.ctrl_cd and a.acct_cd =  " & FilterVar(lgtxtAcctCd , "''", "S") & ""

		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			no_mgnt2 = "TRUE"
			UNIValue(4,0) = " "						
		Else	
			If lgStrflag = 1 then
				stbl_id2 = Trim(lgObjRs("tbl_id"))
				scol_id2 = Trim(lgObjRs("DATA_COLM_ID"))
				temp_nm2 = Trim(lgObjRs("data_colm_nm"))
				sMajor_cd2 = Trim(lgObjRs("MAJOR_CD"))

				UNIValue(4,0) = " select distinct(a.mgnt_val2),b."&temp_nm2 & " from a_open_acct a , "&stbl_id2 & " b"
				UNIValue(4,0) = UNIValue(4,0) & " where a.acct_cd =  " & FilterVar(lgtxtAcctCd , "''", "S") & ""
				UNIValue(4,0) = UNIValue(4,0) & "  and  a.mgnt_val2 = b."&scol_id2
				UNIValue(4,0) = UNIValue(4,0) & "  and  a.mgnt_val2 =  " & FilterVar(lgtxtMgntCd2 , "''", "S") & ""
                                If sMajor_cd2 <> "" then
					UNIValue(4,0) = UNIValue(4,0) & "  and  b.major_cd =  " & FilterVar(sMajor_cd2 , "''", "S") & ""
                                End If  
			Else
				skip_rs5 = "TRUE"	
				UNIValue(4,0) = " "
			End if	
		End If

		Call SubCloseRs(lgObjRs)															'☜: Release RecordSSet	
	Else
		skip_rs5 = "TRUE"		
		UNIValue(4,0) = " "			
	End If				
	
	iArrTemp = split(lgtxtGlNoSeq, gRowSep)		
	
	For ii = 0 To Ubound(iArrTemp,1) - 1
		iArrTemp2 = split(iArrTemp(ii),gColSep)			
		
		If Trim(iArrTemp2(0)) <> "" And Trim(iArrTemp2(1)) <> "" Then
			iStrWhere = iStrWhere & " (a.gl_no <>  " & FilterVar(iArrTemp2(0), "''", "S") & " or a.gl_seq <> " & iArrTemp2(1) & ") and "			 
		End If		
	Next
	
	If InStr(1,iStrWhere, "and") > 0 Then
		iStrWhere = Mid(iStrWhere,1,InStrRev(iStrWhere, "and") -1)	
		iStrWhere = "and	( " & iStrWhere & " ) "
	End If
	
	If iStrWhere = "" Then
		UNIValue(0,8)	= ""
	Else
		UNIValue(0,8)	= iStrWhere
	End If

	UNIValue(0,9)	= lgTailList
      
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    
    Set lgADF = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4)

    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And lgtxtAcctCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtAcctCd_Alt")
		End If
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtAcctNm.value = ""
		End With
		</Script>
<%
    Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtAcctCd.value = "<%=Trim(ConvSPChars(rs1("acct_cd")))%>"
			.txtAcctNm.value = "<%=Trim(ConvSPChars(rs1("acct_nm")))%>"
		End With
		</Script>
<%
    End If
	Set rs1 = Nothing     

    If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And lgtxtGlNo <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtGlNo_Alt")
		End If
    Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtGlNo.value = "<%=Trim(ConvSPChars(rs2("gl_no")))%>"
		End With
		</Script>
<%
    End If
	Set rs2 = Nothing         
    
	If UCase(Trim(skip_rs4)) = "TRUE" Then

	Else    
		If no_mgnt1 = "TRUE" Then											'미결 계정코드에 미결관리코드1이 등록되어 있지 않는 경우 
			If strMsgCd = "" And lgtxtMgntCd1 <> "" Then
				strMsgCd = "970000"		'Not Found
				strMsg1 = Request("txtMgntCd1_Alt")
			End If
	%>
			<Script Language=vbScript>
			With parent.frm1
				.txtMgntCd1Nm.value = ""
			End With
			</Script>
	<%		
		Else															'입력한 미결코드1로 발생된 데이타가 없는 경우 
		    If (rs3.EOF And rs3.BOF) Then
				If strMsgCd = "" And lgtxtMgntCd1 <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtMgntCd1_Alt")
				End If
	%>
				<Script Language=vbScript>
				With parent.frm1
					.txtMgntCd1Nm.value = ""
				End With
				</Script>
	<%
		    Else
	%>
				<Script Language=vbScript>
				With parent.frm1
					.txtMgntCd1.value = "<%=Trim(ConvSPChars(rs3(0)))%>"
					.txtMgntCd1Nm.value = "<%=Trim(ConvSPChars(rs3(1)))%>"
				End With
				</Script>
	<%
		    End If    
			Set rs3 = Nothing
		End If	
	End If

	If UCase(Trim(skip_rs5)) = "TRUE" Then

	Else
		If no_mgnt2 = "TRUE" Then											'미결 계정코드에 미결관리코드2가 등록되어 있지 않는 경우												
			If strMsgCd = "" And lgtxtMgntCd2 <> "" Then
				strMsgCd = "970000"		'Not Found
				strMsg1 = Request("txtMgntCd2_Alt")
			End If
	%>
			<Script Language=vbScript>
			With parent.frm1
				.txtMgntCd2Nm.value = ""
			End With
			</Script>
	<%		
		Else															'입력한 미결코드2로 발생된 데이타가 없는 경우 
		    If (rs4.EOF And rs4.BOF) Then
				If strMsgCd = "" And lgtxtMgntCd2 <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtMgntCd2_Alt")
				End If
	%>
				<Script Language=vbScript>
				With parent.frm1
					.txtMgntCd2Nm.value = ""
				End With
				</Script>
	<%
		    Else
	%>
				<Script Language=vbScript>
				With parent.frm1
					.txtMgntCd2.value = "<%=Trim(ConvSPChars(rs4(0)))%>"
					.txtMgntCd2Nm.value = "<%=Trim(ConvSPChars(rs4(1)))%>"
				End With
				</Script>
	<%
		    End If    
			Set rs4 = Nothing	
		End If
	End If		
	
    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
		Exit Sub
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
' MakeSpreadSheetData
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim  iCtrl_cd
    Dim  iCtrl_Val
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If Len(Trim(lgPageNo)) Then                                        '☜ : Chnage Nextkey str into int value
		If Isnumeric(lgPageNo) Then
			Response.Write lgpageno
			lgPageNo = CInt(lgPageNo)
		End If   
    Else   
		lgPageNo = 0
    End If      
    'rs0에 대한 결과 
    rs0.PageSize     = lgMaxCount                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""     
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			If  Colcnt = 2 Or Colcnt = 5  Then 
				If Trim(rs0(Colcnt-2)) <> "" Then
					Call SubCreateCommandObject(lgObjComm)
					
					iCtrl_cd = Trim(rs0(ColCnt-2))
					iCtrl_val = Trim(rs0(ColCnt-1))
					
					If iCtrl_val <> ""  Then 
						With lgObjComm
						    .CommandText = "USP_A_MGNT_NAME"
						    .CommandType = adCmdStoredProc
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_CD"  ,adVarWChar,adParamInput,Len(iCtrl_cd),iCtrl_cd)
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@INCTRL_VAL" ,adVarWChar,adParamInput,Len(iCtrl_val),iCtrl_val)
						    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@OUT_DATA_COLM_NM",adVarWChar,adParamOutput ,128)
							    
						    lgObjComm.Execute ,, adExecuteNoRecords
						End With
						
						If  Err.number = 0 Then
							iRowStr = iRowStr & Chr(11) & lgObjComm.Parameters("@OUT_DATA_COLM_NM").Value
						End If
					Else
						iRowStr = iRowStr & Chr(11) & ""
					End If
					
					Call SubCloseCommandObject(lgObjComm)
				Else
					iRowStr = iRowStr & Chr(11) & ""
				End If					    
			Else		
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))'rs0(ColCnt)'
			End If				
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
        lgPageNo = ""															'☜: 다음 데이타 없다.
    End If  	
End Sub
%>
<Script Language=vbscript> 
If "<%=lgDataExist%>" = "Yes" Then		
   'Set condition data to hidden area
	With parent
		If "<%=lgPageNo%>" = "1" Or "<%=lgPageNo%>" = ""  Then					' "1" means that this query is first and next data exists
			.Frm1.htxtAcctCd.Value		= .Frm1.txtAcctCd.value
			.Frm1.htxtFromDt.Value		= .Frm1.txtFromDt.text
			.Frm1.htxtToDt.Value		= .Frm1.txtToDt.text
			.Frm1.htxtDocCur.Value		= .Frm1.txtDocCur.value
			.Frm1.htxtBizCd.value    	= .Frm1.txtBizCd.value			
			.Frm1.htxtGlNo.Value		= .Frm1.txtGlNo.Value
			.Frm1.htxtMgntCd1.Value     = .Frm1.txtMgntCd1.Value
			.Frm1.htxtMgntCd2.Value     = .Frm1.txtMgntCd2.Value					
		End If
       
		'Show multi spreadsheet data from this line       
		.ggoSpread.Source	= .frm1.vspdData      
		.ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"						'☜ : Display data
		.lgPageNo			=  "<%=lgPageNo%>"									'☜ : Next next data tag
		
		.DbQueryOk
	End With
End If
</Script>	
