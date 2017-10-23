<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : s2111rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 판매계획등록을 위한 기초자료생성 (Business Logic Asp)						*
'*  6. Comproxy List        : PS2G141.dll
'*  7. Modified date(First) : 2001/01/04																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : Mr. Cho																	*
'* 10. Modifier (Last)      : sonbumyeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/01/04 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<%
    
Dim lgOpModeCRUD
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call LoadBasisGlobalInf()

Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
	
Call SubBizQuery()
Call SubBizQueryMulti()


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Dim ADF												'ActiveX Data Factory 지정 변수선언 
	Dim strRetMsg										'Record Set Return Message 변수선언 
	Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
	Dim rs0												'DBAgent Parameter 선언 
	
	Dim strConPlanNum

	strConPlanNum = Trim(Request("cboConPlanNum"))
		
	
	IF Trim(strConPlanNum) = "" Then
	ELSE
	
		Redim UNISqlId(0)
		Redim UNIValue(0, 0)

		UNISqlId(0) = "ConPlanNum"	
	
		UNIValue(0, 0) = strConPlanNum

		UNILock = DISCONNREAD :	UNIFlag = "1"
	
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

   		' ConPlanNum 명 Display      
			
		If (rs0.EOF And rs0.BOF) Then
			Call DisplayMsgBox("202299", vbOKOnly, "", "", I_MKSCRIPT)
			rs0.Close
			Set rs0 = Nothing
			Set ADF = Nothing
			%>
			<Script Language=vbscript>
				parent.cboConPlanNum.value = ""
				parent.cboConPlanNumNm.value = ""
				parent.cboConPlanNum.Focus()
			</Script>	
			<%
			Response.End
		Else
			%>
			<Script Language=vbscript>
				parent.cboConPlanNumNm.value = "<%=ConvSPChars(rs0("MINOR_NM"))%>"
			</Script>	
			<%    	
			rs0.Close
			Set rs0 = Nothing
			Set ADF = Nothing	
		End If   
	
	END IF
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub    

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
	
	Dim strTemp
	Dim GroupCount
	
    Dim pS21141    
    '-----------------------------------------------
    ' Declare User Variable
    '-----------------------------------------------
    ' ITEM / ITEM GROUP / CUSTOMER 

    Dim i1_b_sales_org1
    Dim i2_s_item_sales_plan1
    ReDim i2_s_item_sales_plan1(S240_I2_sales_grp)
    
    Dim i3_m_timestamp1
    Dim i4_m_timestamp1
    Dim i5_ief_supplied1
    Dim i6_b_item_group1
    Dim i7_b_biz_partner1
    Dim i8_b_item1
    Dim i9_ief_supplied1
    
    Dim i10_s_item_group_sales_plan1
    ReDim i10_s_item_group_sales_plan1(S239_I5_sales_grp)

    ' i2_s_item_sales_plan1    
    Const S240_I2_sp_year = 0
    Const S240_I2_plan_flag = 1
    Const S240_I2_plan_seq = 2
    Const S240_I2_export_flag = 3
    Const S240_I2_cur = 4
    Const S240_I2_sales_grp = 5
    
    ' i_s_item_group_sales_plan1
    Const S239_I5_sales_org = 0
    Const S239_I5_plan_flag = 1
    Const S239_I5_export_flag = 2
    Const S239_I5_cur = 3
    Const S239_I5_plan_seq = 4
    Const S239_I5_sp_year = 5
    Const S239_I5_sales_grp = 6
    
    ' Next Page Variable

    ' Reruen Call Variable
    Dim i1_b_sales_org
    Dim i2_s_item_sales_plan
    Dim i3_m_timestamp
    Dim i4_m_timestamp
    Dim i5_ief_supplied
    Dim i6_b_item_group
    Dim i7_b_biz_partner
    Dim i8_b_item
    Dim i9_ief_supplied
    Dim i10_s_item_group_sales_plan

    ' Export Variables
    Dim exp_b_biz_partner
    Dim exp_b_item_group
    Dim exp_b_item
	Dim E4_b_sales_org
	Dim E5_plan_flag
	Dim E6_export_flag		
    Dim exp_grp   
    
    Dim intGroupCount
    Dim StrNextKey  	
    Dim arrValue
    
    Const C_SHEETMAXROWS_D  = 100
    
    ' exp_grp 저장 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    i3_m_timestamp1 = Trim(Request("txtConSpYearFrom"))
    i4_m_timestamp1 = Trim(Request("txtConSpYearTo"))
    i5_ief_supplied1 = Trim(Request("txtBasicInfo"))
    i9_ief_supplied1 = Trim(Request("txtInfo"))
 
    
    ' 품목별 / 품목 그룹별 / 거래처별 
	Select Case Trim(Request("txtBasicInfo"))
    	
    	Case "T"
    	    
    	    ' 품목별 
    	    i2_s_item_sales_plan1(S240_I2_sp_year) = Trim(Request("txtConSpYear"))
    	    i2_s_item_sales_plan1(S240_I2_plan_flag) = FilterVar(Trim(Request("txtConPlanTypeCd")), "" , "SNM")
    	    i2_s_item_sales_plan1(S240_I2_export_flag) = FilterVar(Trim(Request("txtConDealTypeCd")), "" , "SNM")
    	    i2_s_item_sales_plan1(S240_I2_cur) = Trim(Request("txtConCurr"))
    	    If Len(Trim(Request("cboConPlanNum"))) Then i2_s_item_sales_plan1(S240_I2_plan_seq) = Trim(Request("cboConPlanNum")) 
    	    
	        '---영업조직/영업그룹 구분----
	        Select Case Trim(Request("txtSalesTitle"))    
                
                Case "ORG"
                    i1_b_sales_org1 = FilterVar(Trim(Request("txtConSalesOrg")), "" , "SNM")
                    i2_s_item_sales_plan1(S240_I2_sales_grp) = "S"
                Case "GRP"
                    i1_b_sales_org1 = "S"
                    i2_s_item_sales_plan1(S240_I2_sales_grp) = FilterVar(Trim(Request("txtConSalesOrg")), "" , "SNM")
                    
            End Select
            
    	Case "G"

    	    ' 품목 그룹별 
            i10_s_item_group_sales_plan1(S239_I5_plan_flag) = FilterVar(Trim(Request("txtConPlanTypeCd")), "" , "SNM")
            i10_s_item_group_sales_plan1(S239_I5_export_flag) =  FilterVar(Trim(Request("txtConDealTypeCd")), "" , "SNM")
            i10_s_item_group_sales_plan1(S239_I5_cur) = Trim(Request("txtConCurr"))
		    If Len(Trim(Request("cboConPlanNum"))) Then i10_s_item_group_sales_plan1(S239_I5_plan_seq) = Trim(Request("cboConPlanNum"))
            i10_s_item_group_sales_plan1(S239_I5_sp_year) = Trim(Request("txtConSpYear"))

	        '---영업조직/영업그룹 구분----
	        Select Case Trim(Request("txtSalesTitle"))    
                Case "ORG"
                    i10_s_item_group_sales_plan1(S239_I5_sales_org) = FilterVar(Trim(Request("txtConSalesOrg")), "" , "SNM")
                    i10_s_item_group_sales_plan1(S239_I5_sales_grp) = "S"
                Case "GRP"
                    i10_s_item_group_sales_plan1(S239_I5_sales_org) = "S"
                    i10_s_item_group_sales_plan1(S239_I5_sales_grp) = FilterVar(Trim(Request("txtConSalesOrg")), "" , "SNM")                   
            End Select   

    	Case "C"
    	    
    	    i2_s_item_sales_plan1(S240_I2_sp_year) = Trim(Request("txtConSpYear"))
    	    i2_s_item_sales_plan1(S240_I2_plan_flag) = FilterVar(Trim(Request("txtConPlanTypeCd")), "" , "SNM") 
    	    i2_s_item_sales_plan1(S240_I2_export_flag) = FilterVar(Trim(Request("txtConDealTypeCd")), "" , "SNM") 
    	    i2_s_item_sales_plan1(S240_I2_cur) = Trim(Request("txtConCurr"))
    	    If Len(Trim(Request("cboConPlanNum"))) Then i2_s_item_sales_plan1(S240_I2_plan_seq) = Trim(Request("cboConPlanNum")) 
    	    
	        '---영업조직/영업그룹 구분----
	        Select Case Trim(Request("txtSalesTitle"))    
                
                Case "ORG"
                    i1_b_sales_org1 = FilterVar(Trim(Request("txtConSalesOrg")), "" , "SNM") 
                    i2_s_item_sales_plan1(S240_I2_sales_grp) = "S"
                Case "GRP"
                    i1_b_sales_org1 = "S"
                    i2_s_item_sales_plan1(S240_I2_sales_grp) = FilterVar(Trim(Request("txtConSalesOrg")), "" , "SNM") 
                    
            End Select
            
	End Select 

    i1_b_sales_org              = i1_b_sales_org1              
    i2_s_item_sales_plan        = i2_s_item_sales_plan1        
    i3_m_timestamp              = i3_m_timestamp1              
    i4_m_timestamp              = i4_m_timestamp1              
    i5_ief_supplied             = i5_ief_supplied1             
    i6_b_item_group             = i6_b_item_group1             
    i7_b_biz_partner            = i7_b_biz_partner1            
    i8_b_item                   = i8_b_item1                   
    i9_ief_supplied             = i9_ief_supplied1             
    i10_s_item_group_sales_plan = i10_s_item_group_sales_plan1 

	iStrPrevKey = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	
	
	If iStrPrevKey <> "" then					
	    
		arrValue = Split(iStrPrevKey, gColSep)
        
	else			

	End If
    
	set pS21141 = Server.CreateObject ("PS2G141.cSCustSPSvr")
	
	' Call the Dll   		
    Call pS21141.S_SALES_PLAN_BASIC_DATA ( gStrGlobalCollection, _
                           i1_b_sales_org, i2_s_item_sales_plan, i3_m_timestamp, i4_m_timestamp, _
                           i5_ief_supplied, i6_b_item_group, i7_b_biz_partner, i8_b_item, _
                           i9_ief_supplied, i10_s_item_group_sales_plan, _
                           exp_b_biz_partner, exp_b_item_group, exp_b_item, _
                           E4_b_sales_org, E5_plan_flag, E6_export_flag, exp_grp )

	If CheckSYSTEMError(Err,TRUE) = True Then

		Response.Write "<Script language=vbs>  " & vbCr   
		Response.Write " Parent.txtConSalesOrgNm.value  = """ & ConvSPChars(E4_b_sales_org) & """" & vbCr    
		Response.Write " Parent.txtConPlanTypeNm.value  = """ & ConvSPChars(E5_plan_flag)	& """" & vbCr    
		Response.Write " Parent.txtConDealTypeNm.value  = """ & ConvSPChars(E6_export_flag) & """" & vbCr    
		Response.Write " Parent.txtConSalesOrg.focus  " & vbCr   
	    Response.Write "</Script>      " & vbCr      

       Set pS21141 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   
    
    Set pS21141 = Nothing	

    iLngMaxRow  = CLng(Request("txtMaxRows"))										'☜: Fetechd Count      

    For iLngRow = 0 To UBound(exp_grp,1)

		Select Case Trim(Request("txtBasicInfo"))
		Case "T"    ' 품목별 
			istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, 29))
			istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, 30))					
            ' 계획 단위 
			istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, 0))
		Case "G"    ' 품목그룹별 
			istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, 25))
			istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, 26))						
			' 계획단위 
			istrdata = istrdata & Chr(11) & ""
		Case "C"    ' 거래처별 
			istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, 27))
			istrdata = istrdata & Chr(11) & ConvSPChars(exp_grp(iLngRow, 28))					
			' 계획단위 
			istrdata = istrdata & Chr(11) & ""
		End Select        
        
        ' 합계란 
        istrdata = istrdata & Chr(11) & 0
        istrdata = istrdata & Chr(11) & 0

        ' 월별 수량/금액 
		Dim iCnt
        For iCnt = 1 to 12        
			istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, iCnt), ggAmtOfMoney.DecPoint, 0)
			istrdata = istrdata & Chr(11) & UNINumClientFormat(exp_grp(iLngRow, iCnt + 12), ggAmtOfMoney.DecPoint, 0)
		Next

        'istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12) 
        
    Next

    Response.Write "<Script language=vbs> " & vbCr   
    
    Response.Write " Parent.txtConSalesOrgNm.value  = """ & ConvSPChars(E4_b_sales_org) & """" & vbCr    
    Response.Write " Parent.txtConPlanTypeNm.value  = """ & ConvSPChars(E5_plan_flag)	& """" & vbCr    
    Response.Write " Parent.txtConDealTypeNm.value  = """ & ConvSPChars(E6_export_flag) & """" & vbCr    
            
    Response.Write " Parent.HConSalesOrg.value		= """ & Trim(Request("txtConSalesOrg"))		& """" & vbCr    
    Response.Write " Parent.HConSpYear.value		= """ & Trim(Request("txtConSpYear"))		& """" & vbCr    
    Response.Write " Parent.HPlanTypeCd.value		= """ & Trim(Request("txtConPlanTypeCd"))	& """" & vbCr    
    Response.Write " Parent.HConDealTypeCd.value	= """ & Trim(Request("txtConDealTypeCd"))	& """" & vbCr    
    Response.Write " Parent.HConCurr.value			= """ & Trim(Request("txtConCurr"))			& """" & vbCr    
    Response.Write " Parent.HConPlanNum.value		= """ & Trim(Request("cboConPlanNum"))		& """" & vbCr    
    Response.Write " Parent.HConFrmYear.value		= """ & Trim(Request("txtConSpYearFrom"))	& """" & vbCr    
    Response.Write " Parent.HConToYear.value		= """ & Trim(Request("txtConSpYearTo"))		& """" & vbCr 
    Response.Write " Parent.HBasicInfo.value		= """ & Trim(Request("txtBasicInfo"))		& """" & vbCr    
    Response.Write " Parent.HInfo.value				= """ & Trim(Request("txtInfo"))			& """" & vbCr 
    Response.Write " Parent.ggoSpread.Source        = Parent.vspdData								 " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip      """ & istrData							& """" & vbCr
    Response.Write " Parent.lgStrPrevKey            = """ & StrNextKey							& """" & vbCr  
    Response.Write " Parent.DbQueryOk	" & vbCr   
    Response.Write "</Script>	"		  & vbCr      
                   	
End Sub    


'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>
