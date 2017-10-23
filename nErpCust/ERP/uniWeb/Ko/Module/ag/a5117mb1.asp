<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","QB")
Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QA")

Err.Clear
On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4 ,rs5,rs6     '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim lgtxtFromGlDt
Dim lgtxtUsr_Id
Dim lgtxtToGlDt
Dim lgtxtBizArea
Dim lgtxtBizArea1
Dim lgtxtCOST_CENTER_CD
Dim lgtxtdeptcd
Dim lgcboGlInputType
Dim lgcdoConfig
Dim lgtxtMaxRows

Dim dr_loc_amt
Dim cr_loc_amt
Dim biz_area_nm
Dim biz_area_nm1
Dim Usr_Nm
Dim cost_nm
Dim dept_nm

Dim StrDesc, StrRefNo,strAmtFr, strAmtTo

Dim strSql

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = Request("lgPageNo")                               '�� : Next key flag
 '   lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    
	' ���Ѱ��� �߰� 
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
Sub TrimData()

	lgtxtFromGlDt		= UNIConvDate(Request("txtFromGlDt"))
	lgtxtToGlDt			= UNIConvDate(Request("txtToGlDt"))
	lgtxtBizArea		= Request("txtBizArea")
	lgtxtBizArea1		= Request("txtBizArea1")	
	lgtxtCOST_CENTER_CD	= Request("txtCOST_CENTER_CD")
	lgtxtdeptcd			= Request("txtdeptcd")
	lgcboGlInputType	= Trim(Request("cboGlInputType"))
	lgcboConfFg			= Trim(Request("cboConfFg"))
	lgtxtUsr_Id			= Request("txtUsr_Id")
	lgtxtMaxRows		= Request("txtMaxRows")

	StrDesc				= Request("txtDesc")
	StrRefNo			= Request("txtRefNo")
	strAmtFr			= UNIConvNum(Request("txtAmtFr"),0)
	strAmtTo			= UNIConvNum(Request("txtAmtTo"),0)
	
	strSql = " "
	
	IF StrRefNo <> "" then
		strSql  = strSql + "  AND A.REF_NO  LIKE   " & FilterVar(StrRefNo & "%", "''", "S") & " "
	end if

	IF StrDesc <> "" then
		strSql  = strSql + "  AND A.TEMP_GL_DESC LIKE  " & FilterVar("%" & StrDesc & "%", "''", "S") & " "
	end if

'-------------------------
'�ݾ� 
'-------------------------

	If strAmtFr <> 0 or strAmtTo <> 0 Then
		If strAmtFr > 0 and strAmtTo <= 0 Then
			strSql = strSql & " AND (a.DR_LOC_AMT >= " & strAmtFr & " AND  a.CR_LOC_AMT >= " & strAmtFr & " ) "
		ElseIf strAmtFr <= 0 and strAmtTo > 0 Then
			strSql = strSql & " AND (a.DR_LOC_AMT <= " & strAmtTo & " AND  a.CR_LOC_AMT <= " & strAmtTo & " ) "
		Else
			strSql = strSql & " AND (a.DR_LOC_AMT between " & strAmtFr & " AND " & strAmtTo & " AND a.CR_LOC_AMT between  " & strAmtFr & " AND " & strAmtTo & " ) "
		End If
	End If
	
	if lgtxtBizArea = "" then
		strSql = strSql & " and A.biz_area_cd >= " & FilterVar("0", "''", "S") & " "
	else		
		strSql = strSql & " and A.biz_area_cd >=  " & FilterVar(lgtxtBizArea , "''", "S") & ""
	end if
	
	if lgtxtBizArea1 = "" then
		strSql = strSql & " and A.biz_area_cd <= " & FilterVar("ZZZZZZZZZZZ", "''", "S") & " "
	else		
		strSql = strSql & " and A.biz_area_cd <=  " & FilterVar(lgtxtBizArea1 , "''", "S") & ""
	end if	

	IF lgtxtCOST_CENTER_CD <> "" then
		strSql  = strSql +  " AND A.COST_CD =  " & FilterVar(lgtxtCOST_CENTER_CD , "''", "S") & " "
	end if

	IF lgtxtdeptcd <> "" then
		strSql  = strSql + " AND A.DEPT_CD =  " & FilterVar(lgtxtdeptcd , "''", "S") & " "
	end if

	IF lgcboGlInputType <> "" then
		strSql  = strSql + " AND A.GL_INPUT_TYPE =  " & FilterVar(lgcboGlInputType , "''", "S") & " "
	end if
	
	IF lgcboConfFg	 <> "" then
		strSql  = strSql + " AND A.conf_fg =  " & FilterVar(lgcboConfFg , "''", "S") & " "
	end if

	IF lgtxtdeptcd <> "" then
		strSql  = strSql +  " AND C.ORG_CHANGE_ID  = " & " " & FilterVar(request("OrgChangeId"), "''", "S") & "  "
	end if

	IF lgtxtUsr_Id <> "" then
		strSql  = strSql +  " AND A.INSRT_USER_ID  = " & " " & FilterVar(lgtxtUsr_Id, "''", "S") & "  "
	end if
	
	iF Request("lgAuthorityFlag") = "Y" then      '���Ѱ��� �߰� 
		strSql = strSql & " and EXISTS ( SELECT 1 FROM z_usr_authority_value S WHERE a.dept_cd = S.code_value and S.usr_id =  " & FilterVar(gUsrID , "''", "S") & " AND S.module_cd = " & FilterVar("A", "''", "S") & "  )  "   '���Ѱ��� �߰� 
	end if

	strSql = strSql + " "

	' ���Ѱ��� �߰� 
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

	strSql	= strSql	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

	strSql = strSql + " "


End Sub

Sub MakeSpreadSheetData()

Const C_SHEETMAXROWS_D  = 100                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 

    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If Len(Trim(lgPageNo)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          lgPageNo = CInt(lgPageNo)
       End If   
    Else   
       lgPageNo = 0
    End If   
    'rs0�� ���� ��� 
    rs0.PageSize     = C_SHEETMAXROWS_D                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1

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

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
    
    'rs1�� ���� ��� 
    IF NOT (rs1.EOF or rs1.BOF) then
		dr_loc_amt = rs1("dr_loc_amt")
		cr_loc_amt = rs1("cr_loc_amt")
    End if
    rs1.Close
    Set rs1 = Nothing 
    
    'rs2�� ���� ��� 
    IF NOT (rs2.EOF or rs2.BOF) then
	    biz_area_nm = rs2("biz_area_nm")
    END IF
    rs2.Close
    Set rs2 = Nothing
    
    'rs3�� ���� ���    
    IF NOT (rs3.EOF or rs3.BOF) then
		cost_nm = rs3("cost_nm")
    END IF
    rs3.Close
    Set rs3 = Nothing
    
    'rs4�� ���� ��� 
    IF NOT (rs4.EOF or rs4.BOF) then
		dept_nm = rs4("dept_nm")
    END IF
    rs4.Close
    Set rs4 = Nothing
    
    'rs5�� ���� ��� 
    IF NOT (rs5.EOF or rs5.BOF) then
	    biz_area_nm1 = rs5("biz_area_nm")
    END IF
    rs5.Close
    Set rs5 = Nothing    
    
    IF NOT (rs6.EOF or rs6.BOF) then
	    Usr_Nm = rs6("Usr_Nm")
    END IF
    rs6.Close
    Set rs6 = Nothing    

End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(6)                                                    '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(6,9)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "a5117ma101"
	UNISqlId(1) = "A5117MA102"
    UNISqlId(2) = "ABIZNM"
    UNISqlId(3) = "M6111QA104"
    UNISqlId(4) = "ADEPTNM"
    UNISqlId(5) = "ABIZNM"
    UNISqlId(6) = "CommonQry"    
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    
    'rs0�� ���� Value�� setting    
    UNIValue(0,0)  = lgSelectList  
  	UNIValue(0,1)  = FilterVar(lgtxtFromGlDt,"","S")	'UNIConvDate(Request("txtFromGlDt") )
	UNIValue(0,2)  = FilterVar(lgtxtToGlDt,"","S")
	UNIValue(0,3)  = strSql
	   'Call SvrMsgBox(lgSelectList , vbInformation, I_MKSCRIPT)

	'rs1�� ���� Value�� setting
  	
	UNIValue(1,0)  = FilterVar(lgtxtFromGlDt,"","S")
	UNIValue(1,1)  = FilterVar(lgtxtToGlDt,"","S")
	UNIValue(1,2)  = strSql

    
    'rs2�� ���� Value�� setting
	UNIValue(2,0) = " " & FilterVar(lgtxtBizArea, "''", "S") & ""
	
	'rs3�� ���� Value�� setting
	UNIValue(3,0)  = FilterVar(lgtxtCOST_CENTER_CD , "''", "S")				                           '�Էµ� ���� ������ ���̰��� �Ѱ��ش� 
		
	'rs4�� ���� Value�� setting
	UNIValue(4,0)  = FilterVar(lgtxtdeptcd , "''", "S")	
	UNIValue(4,1)  = FilterVar(request("OrgChangeId"), "''", "S")
	
	UNIValue(5,0) = " " & FilterVar(lgtxtBizArea1, "''", "S") & ""	
	
	UNIValue(6,0) = " select USR_ID,USR_NM from Z_USR_MAST_REC where USR_ID = " & FilterVar(UCase(lgtxtUsr_Id),"''","S")

    '--------------- ������ coding part(�������,End)------------------------------------------------------
       
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

	On Error Resume Next

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
	
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4,rs5,rs6)	
    Set lgADF = Nothing                                                    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMsgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
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

%>

<Script Language=vbscript>
 
    If "<%=lgDataExist%>" = "Yes" Then

       
       With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
					.Frm1.htxtFromGlDt.Value		= .Frm1.txtFromGlDt.text
					.Frm1.htxtToGlDt.Value			= .Frm1.txtToGlDt.text
					.Frm1.htxtBizArea.Value			= .Frm1.txtBizArea.Value
					.Frm1.htxtBizArea1.Value		= .Frm1.txtBizArea1.Value					
					.Frm1.htxtCOST_CENTER_CD.Value  = .Frm1.txtCOST_CENTER_CD.Value
					.Frm1.htxtdeptcd.Value			= .Frm1.txtdeptcd.Value
					.Frm1.hcboGlInputType.Value     = .Frm1.cboGlInputType.Value
					.Frm1.htxtDesc.Value			= .Frm1.txtDesc.Value
					.Frm1.htxtRefNo.Value			= .Frm1.txtRefNo.Value
					.Frm1.htxtAmtFr.Value			= .Frm1.txtAmtFr.Value
					.Frm1.htxtAmtTo.Value			= .Frm1.txtAmtTo.Value
					.Frm1.hcboConfFg.Value			= .Frm1.cboConfFg.Value
					.Frm1.htxtUsr_Id.Value			= .Frm1.txtUsr_Id.Value
			End If
       
        'Show multi spreadsheet data from this line       
        .ggoSpread.Source	= .frm1.vspdData      
        .ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
        .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag
       
       																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		.frm1.txtDrlocAmt.text			= "<%=UNINumClientFormat(dr_loc_amt, ggAmtOfMoney.DecPoint, 0)%>"		
		.frm1.txtCrlocAmt.text			= "<%=UNINumClientFormat(cr_loc_amt, ggAmtOfMoney.DecPoint, 0)%>"		
		.frm1.txtBizAreaNm.value		= "<%=biz_area_nm%>"
		.frm1.txtBizAreaNm1.value		= "<%=biz_area_nm1%>"		
		.frm1.txtCOST_CENTER_NM.value	= "<%=cost_nm%>"
		.frm1.txtdeptnm.value			= "<%=dept_nm%>"
		.frm1.txtUsr_NM.value			= "<%=USR_NM%>"
		
	   End With
       
       
       Parent.DbQueryOk
    Else
		With parent
		.frm1.txtDrlocAmt.text			= "<%=UNINumClientFormat(dr_loc_amt, ggAmtOfMoney.DecPoint, 0)%>"										  
		.frm1.txtCrlocAmt.text			= "<%=UNINumClientFormat(cr_loc_amt, ggAmtOfMoney.DecPoint, 0)%>"		
		.frm1.txtBizArea.value			= ""
		.frm1.txtBizAreaNm.value		= ""
		.frm1.txtBizArea1.value			= ""
		.frm1.txtBizAreaNm1.value		= ""		
		.frm1.txtCOST_CENTER_Cd.value	= ""
		.frm1.txtCOST_CENTER_NM.value	= ""
		.frm1.txtdeptCd.value			= ""
		.frm1.txtdeptnm.value			= ""
		.frm1.cboGlInputType.Value		= ""
		.Frm1.txtDesc.Value				= ""
		.Frm1.txtRefNo.Value			= ""
		.Frm1.txtAmtFr.Text				= ""
		.Frm1.txtAmtTo.Text				= ""
		.Frm1.cboConfFg.Value			= ""
		.Frm1.txtUsr_Id.Value			= ""
		.Frm1.txtUsr_NM.Value			= ""
		
		End With
	End if

</Script>	


