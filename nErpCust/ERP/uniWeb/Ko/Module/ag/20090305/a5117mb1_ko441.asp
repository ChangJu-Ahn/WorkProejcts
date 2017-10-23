
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% 

    Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
    Call LoadBasisGlobalInf()

    Dim lgStrPrevKey
	Dim lgtxtFromGlDt
	Dim lgtxtUsr_Id
	Dim lgtxtToGlDt
	Dim lgtxtBizArea
	Dim lgtxtBizArea1
	Dim lgtxtCOST_CENTER_CD
	Dim lgtxtdeptcd
	Dim lgcboGlInputType
	Dim lgcboConfFg
	Dim lgtxtMaxRows
	
	Dim dr_loc_amt
	Dim cr_loc_amt
	Dim dr_cr_loc_amt
	
	Dim biz_area_nm
	Dim biz_area_nm1
	Dim Usr_Nm
	Dim cost_nm
	Dim dept_nm
	

	Dim lgDataExist
    'Dim lgstrData
	
	Dim StrDesc, StrRefNo,strAmtFr, strAmtTo

    On Error Resume Next                                                   '☜: Protect prorgram from crashing

    Err.Clear                                                              '☜: Clear Error status
    
    Call HideStatusWnd                                                     '☜: Hide Processing message

    lgErrorStatus  = ""
    lgStrPrevKey   = UNICInt(Trim(Request("lgStrPrevKey")), 0)                   '☜: Next Key
   
	lgtxtFromGlDt		= UNIConvDate(Trim(Request("txtFromGlDt")))
	lgtxtToGlDt			= UNIConvDate(Trim(Request("txtToGlDt")))
	lgtxtBizArea		= Trim(Request("txtBizArea"))
	lgtxtBizArea1		= Trim(Request("txtBizArea1"))
	lgtxtCOST_CENTER_CD	= Trim(Request("txtCOST_CENTER_CD"))
	lgtxtdeptcd			= Trim(Request("txtdeptcd"))
	lgcboGlInputType	= Trim(Request("cboGlInputType"))
	lgcboConfFg			= Trim(Request("cboConfFg"))
	lgtxtUsr_Id			= Trim(Request("txtUsr_Id"))
	lgtxtMaxRows		= Request("txtMaxRows")

	StrDesc				= Trim(Request("txtDesc"))
	StrRefNo			= Trim(Request("txtRefNo"))
	strAmtFr			= Trim(Request("txtAmtFr"))
	strAmtTo			= Trim(Request("txtAmtTo"))							   '본지점계정외' 1:미포함, 2:포함
	
	lgDataExist    = "No"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim lgStrSQL
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
                                                  '☜ : Release RecordSSet
    Call SubBizQueryMulti()		'상단 그리드 
    Call SubDrCrTotAmt()		'합계

End Sub    


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim lgStrSQL
    Dim iDx
    Dim iSelCount
  	  
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    lgDataExist    = "Yes"
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status
'Call ServerMesgBox(lgStrPrevKey , vbInformation, I_MKSCRIPT)
	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------

	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT   TOP " & iSelCount & " a.SEQ, a.TEMP_GL_DT,   "
	lgStrSQL = lgStrSQL & " 		 a.BIZ_AREA_CD,      "
	lgStrSQL = lgStrSQL & "          a.BIZ_AREA_NM,      "
	lgStrSQL = lgStrSQL & " 		 a.TEMP_GL_NO,       "
	lgStrSQL = lgStrSQL & " 		 b.GL_NO,            "
	lgStrSQL = lgStrSQL & " 		 INSRT_USER_ID = dbo.ufn_x_getcodename('Z_USR_MAST_REC',b.UPDT_USER_ID,''),    "
	lgStrSQL = lgStrSQL & " 		 a.ACCT_CD,          "
	lgStrSQL = lgStrSQL & " 		 a.ACCT_NM,          "
	lgStrSQL = lgStrSQL & " 		 a.DR_ITEM_LOC_AMT,  "
	lgStrSQL = lgStrSQL & " 		 a.CR_ITEM_LOC_AMT,  "
	lgStrSQL = lgStrSQL & " 		 a.TEMP_GL_DESC,     "
	lgStrSQL = lgStrSQL & " 		 a.GL_INPUT_TYPE,    "
	lgStrSQL = lgStrSQL & "          a.GL_INPUT_TYPE_NM, "
	lgStrSQL = lgStrSQL & " 		 a.DEPT_CD,          "
	lgStrSQL = lgStrSQL & " 		 a.DEPT_NM,          "
	lgStrSQL = lgStrSQL & " 		 a.COST_CD,          "
	lgStrSQL = lgStrSQL & " 		 a.COST_NM,          "		
	lgStrSQL = lgStrSQL & " 		 a.ITEM_SEQ          "
	lgStrSQL = lgStrSQL & " FROM     ufn_a5117ma1_ko441("
	lgStrSQL = lgStrSQL &  		 	 FilterVar(lgtxtFromGlDt,"''", "S") & ", "
	lgStrSQL = lgStrSQL &  		 	 FilterVar(lgtxtToGlDt,"''", "S") & ", "
	lgStrSQL = lgStrSQL &  		 	 FilterVar(StrRefNo & "%", "''", "S") & ", "	
	lgStrSQL = lgStrSQL &  		 	 FilterVar(StrDesc & "%", "''", "S") & ", "	

	if lgtxtBizArea = "" then
		lgStrSQL = lgStrSQL & 		 FilterVar("0", "''", "S") & ", "
	else		
		lgStrSQL = lgStrSQL & 		 FilterVar(lgtxtBizArea , "''", "S") & ", "
	end if
	
	if lgtxtBizArea1 = "" then
		lgStrSQL = lgStrSQL & 		 FilterVar("ZZZZZZZZZZZ", "''", "S") & ", "
	else		
		lgStrSQL = lgStrSQL &		 FilterVar(lgtxtBizArea1 , "''", "S") & ", "
	end if		

	If strAmtFr <> "" Then
		lgStrSQL = lgStrSQL & 		 UNIConvNum(strAmtFr,0) & ", "
	ELSE
		lgStrSQL = lgStrSQL & 		 UNIConvNum("-9999999999999", 0) & ", "
	End If
	
	If strAmtTo <> "" Then		
		lgStrSQL = lgStrSQL & 		 UNIConvNum(strAmtTo,0) & ", "
	ELSE
		lgStrSQL = lgStrSQL & 		 UNIConvNum("9999999999999", 0) & ", "		
	End If	

		lgStrSQL  = lgStrSQL &  	FilterVar(lgtxtCOST_CENTER_CD & "%", "''", "S") & ", "
		lgStrSQL  = lgStrSQL &		FilterVar(lgtxtdeptcd & "%", "''", "S") & ", "
		lgStrSQL  = lgStrSQL & 		FilterVar(lgcboGlInputType & "%", "''", "S") & ", "
		lgStrSQL  = lgStrSQL &		FilterVar(Trim(request("OrgChangeId")) & "%", "''", "S") & ", "	
		lgStrSQL  = lgStrSQL & 		FilterVar(lgcboConfFg & "%", "''", "S") & ", "
	If lgtxtUsr_Id <> "" Then			
		lgStrSQL  = lgStrSQL &  	FilterVar(lgtxtUsr_Id & "", "''", "S") & ") a, "
	ELSE
		lgStrSQL  = lgStrSQL &  	FilterVar(lgtxtUsr_Id & "%", "''", "S") & ") a, "
	End If
		lgStrSQL  = lgStrSQL &  	" a_gl b(nolock) where a.gl_no *= b.gl_no "

    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
'Call ServerMesgBox(lgStrSQL , vbInformation, I_MKSCRIPT)	    	
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
        lgStrPrevKey  = ""
        lgErrorStatus = "YES"
        Exit Sub 
    Else    
	   If CDbl(lgStrPrevKey) > 0 Then
		  lgObjRs.Move     = CDbl(C_SHEETMAXROWS_D) * CDbl(lgStrPrevKey)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	   End If 
	    
       iDx = 1		
       
       lgstrData = ""
       lgLngMaxRow       = CLng(Request("txtMaxRows"))

       
       Do While Not lgObjRs.EOF

          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ"))               	
          lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("TEMP_GL_DT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))      
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEMP_GL_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INSRT_USER_ID"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DR_ITEM_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CR_ITEM_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEMP_GL_DESC"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_INPUT_TYPE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GL_INPUT_TYPE_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COST_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COST_NM"))                    
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_SEQ"))           
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)
'Call ServerMesgBox(lgstrData , vbInformation, I_MKSCRIPT)	
          lgObjRs.MoveNext

          iDx =  iDx + 1
         If iDx > C_SHEETMAXROWS_D Then
			 lgStrPrevKey = lgStrPrevKey + 1
             Exit Do
         End If        
      Loop 
    End If
        
    If iDx <= C_SHEETMAXROWS_D Then
	    lgStrPrevKey = ""            
    End If            
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
       Exit Sub
    End If   

End Sub    



'============================================================================================================
' Name : SubDrCrTotAmt()
' Desc : 차대 합계 금액
'============================================================================================================
Sub SubDrCrTotAmt()
    Dim lgStrSQL
    Dim iDx
   
    
    lgDataExist    = "Yes"
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------

	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT   SUM(DR_ITEM_LOC_AMT) AS TOT_DR_ITEM_LOC_AMT,   "
	lgStrSQL = lgStrSQL & " 		 SUM(CR_ITEM_LOC_AMT) AS TOT_CR_ITEM_LOC_AMT,   "
	lgStrSQL = lgStrSQL & " 		 SUM(DR_ITEM_LOC_AMT) - SUM(CR_ITEM_LOC_AMT) AS DR_CR_ITEM_LOC_AMT   "
	lgStrSQL = lgStrSQL & " FROM     ufn_a5117ma1_ko441("
	lgStrSQL = lgStrSQL &  		 	 FilterVar(lgtxtFromGlDt,"''", "S") & ", "
	lgStrSQL = lgStrSQL &  		 	 FilterVar(lgtxtToGlDt,"''", "S") & ", "
	lgStrSQL = lgStrSQL &  		 	 FilterVar(StrRefNo & "%", "''", "S") & ", "	
	lgStrSQL = lgStrSQL &  		 	 FilterVar(StrDesc & "%", "''", "S") & ", "	

	if lgtxtBizArea = "" then
		lgStrSQL = lgStrSQL & 		 FilterVar("0", "''", "S") & ", "
	else		
		lgStrSQL = lgStrSQL & 		 FilterVar(lgtxtBizArea , "''", "S") & ", "
	end if
	
	if lgtxtBizArea1 = "" then
		lgStrSQL = lgStrSQL & 		 FilterVar("ZZZZZZZZZZZ", "''", "S") & ", "
	else		
		lgStrSQL = lgStrSQL &		 FilterVar(lgtxtBizArea1 , "''", "S") & ", "
	end if		

	If strAmtFr <> "" Then
		lgStrSQL = lgStrSQL & 		 UNIConvNum(strAmtFr,0) & ", "
	ELSE
		lgStrSQL = lgStrSQL & 		 UNIConvNum("-9999999999999", 0) & ", "
	End If
	
	If strAmtTo <> "" Then		
		lgStrSQL = lgStrSQL & 		 UNIConvNum(strAmtTo,0) & ", "
	ELSE
		lgStrSQL = lgStrSQL & 		 UNIConvNum("9999999999999", 0) & ", "		
	End If	

		lgStrSQL  = lgStrSQL &  	FilterVar(lgtxtCOST_CENTER_CD & "%", "''", "S") & ", "
		lgStrSQL  = lgStrSQL &		FilterVar(lgtxtdeptcd & "%", "''", "S") & ", "
		lgStrSQL  = lgStrSQL & 		FilterVar(lgcboGlInputType & "%", "''", "S") & ", "
		lgStrSQL  = lgStrSQL &		FilterVar(Trim(request("OrgChangeId")) & "%", "''", "S") & ", "	
		lgStrSQL  = lgStrSQL & 		FilterVar(lgcboConfFg & "%", "''", "S") & ", "
	If lgtxtUsr_Id <> "" Then			
		lgStrSQL  = lgStrSQL &  	FilterVar(lgtxtUsr_Id & "", "''", "S") & ")  "
	ELSE
		lgStrSQL  = lgStrSQL &  	FilterVar(lgtxtUsr_Id & "%", "''", "S") & ")  "
	End If
	'	lgStrSQL  = lgStrSQL &  	FilterVar(lgtxtUsr_Id & "%", "''", "S") & ")  "
		lgStrSQL  = lgStrSQL & " WHERE SEQ = 2 "
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
    	
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 

        lgErrorStatus = "YES"
        Exit Sub 
    Else    

        dr_loc_amt 		= UNINumClientFormat(lgObjRs("TOT_DR_ITEM_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)
        cr_loc_amt 		= UNINumClientFormat(lgObjRs("TOT_CR_ITEM_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)          
        dr_cr_loc_amt 	= UNINumClientFormat(lgObjRs("DR_CR_ITEM_LOC_AMT"), ggAmtOfMoney.DecPoint, 0)            

    End If
          
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
       Exit Sub
    End If   

End Sub  


%>

<Script Language=vbscript>
 
    If "<%=lgDataExist%>" = "Yes" Then

 'msgbox "x2"
       With parent
			If "<%=lgStrPrevKey%>" = "1" Then   ' "1" means that this query is first and next data exists
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
        .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
 'msgbox "<%=lgstrData%>"        
        .lgStrPrevKey			=  "<%=lgStrPrevKey%>"          '☜ : Next next data tag
       
       																	'☜: 화면 처리 ASP 를 지칭함 
		.frm1.txtNDrAmt.text			= "<%=dr_loc_amt%>"		
		.frm1.txtNCrAmt.text			= "<%=cr_loc_amt%>"		
		.frm1.txtNSumAmt.text			= "<%=dr_cr_loc_amt%>"	

	   End With
       
       
       Parent.DbQueryOk
    else
    	
		With parent
		.frm1.txtNDrAmt.text			= "<%=dr_loc_amt%>"		
		.frm1.txtNCrAmt.text			= "<%=cr_loc_amt%>"		
		.frm1.txtNSumAmt.text			= "<%=dr_cr_loc_amt%>"	
			
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
