<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->


<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2101mb1
'*  4. Program Name         : 예산계정등록 
'*  5. Program Desc         : Register of Budget Account/Accout Group
'*  6. Comproxy List        : FU0011, FU0018
'*  7. Modified date(First) : 2000.09.14
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'=======================================================================================================


'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear 
      
Dim lgOpModeCRUD
    
Const C_CTRL_FG = 0
Const C_BDG_CD = 1
Const C_BDG_NM = 2
Const C_ACCT_CD = 3
Const C_ACCT_NM = 4
Const C_GP_CD = 5
Const C_GP_NM = 6
Const C_CTRL_UNIT = 7
Const C_TRANS_FG =  8
Const C_DIVERT_FG = 9
Const C_ADD_FG = 10	

Call LoadBasisGlobalInf()
	
Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update             
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status
    	
Dim PAFG205LIST		
Dim EG1_export_group
Dim iStrData
	
Dim E2_f_bdg_acct
Const C_E2_bdg_cd = 0      'bdg_cd
Const C_E2_gp_acct_nm = 1       'gp_acct_nm
Const C_E2_acct_ctrl_fg = 2
    
Dim iLngRow,iLngCol
	
'==============신규================
Dim iIntQueryCount
Dim iIntLoopCount
	
Dim I1_f_bdg_acct
Const C_I1_bdg_cd = 0
Const C_I1_acct_ctrl_fg = 1
Const C_I1_next_bdg_cd = 2
    
Dim iStrPrevKey	
Dim txtBDG_NM
Dim iIntMaxRows
'==================================
			    
Const C_SHEETMAXROWS = 100  
	
	iStrPrevKey		= Trim(Request("lgStrPrevKey"))   
	iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    
    Redim I1_f_bdg_acct(C_I1_next_bdg_cd)
    I1_f_bdg_acct(C_I1_bdg_cd) = Request("txtBDG_CD")
    I1_f_bdg_acct(C_I1_acct_ctrl_fg) = Request("cbofg")
    I1_f_bdg_acct(C_I1_next_bdg_cd) = iStrPrevKey        
    
    If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)          
       End If   
    Else   
       iIntQueryCount = 0
    End If
    
	If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
            
    Set PAFG205LIST = Server.CreateObject("PAFG205.cFListBdgAcctSvr")
	
	If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
  
    Call PAFG205LIST.F_LIST_BDG_ACCT_SVR(gStrGlobalCollection,_
										C_SHEETMAXROWS,_
										I1_f_bdg_acct,_
										EG1_export_group,_
										E2_f_bdg_acct)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PAFG205LIST = Nothing
       Exit Sub
    End If    
        
    Set PAFG205LIST = nothing    
	
    iStrData = ""
    iIntLoopCount = 0	

    If isArray(EG1_export_group) Then
		For iLngRow = 0 To UBound(EG1_export_group, 1) 		
			iIntLoopCount = iIntLoopCount + 1
    
		    If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
					iStrData = iStrData & Chr(11) & EG1_export_group(iLngRow, C_CTRL_FG)
					iStrData = iStrData & Chr(11) & ""
					iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, C_BDG_CD))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_BDG_NM)))
					iStrData = iStrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, C_ACCT_CD))
					istrData = iStrData & Chr(11) & ""
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_ACCT_NM)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_GP_CD)))
					iStrData = iStrData & Chr(11) & ""
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_GP_NM)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(EG1_export_group(iLngRow, C_CTRL_UNIT)))
					iStrData = iStrData & Chr(11) & ""
					iStrData = iStrData & Chr(11) & Trim(EG1_export_group(iLngRow, C_TRANS_FG))
					iStrData = iStrData & Chr(11) & Trim(EG1_export_group(iLngRow, C_DIVERT_FG))
					iStrData = iStrData & Chr(11) & Trim(EG1_export_group(iLngRow, C_ADD_FG))
					IStrData = iStrData & Chr(11) & Cstr(iIntMaxRows + iLngRow + 1)
				 
				    iStrData = iStrData & Chr(11) & Chr(12)
		    Else				
				iStrPrevKey = EG1_export_group(UBound(EG1_export_group, 1), C_BDG_CD)
				iIntQueryCount = iIntQueryCount + 1
				Exit For
				  
			End If
		Next
	End If	
	
	If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then	
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If
	
	If IsEmpty(E2_f_bdg_acct) = False Then
		 txtBDG_NM = E2_f_bdg_acct(C_E2_gp_acct_nm)
	End If
	
	Response.Write " <Script Language=vbscript>								 " & vbCr
	Response.Write " With parent											 " & vbCr
    Response.Write "	.ggoSpread.Source		 = .frm1.vspdData			 " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData	   """ & iStrData		& """" & vbCr
    Response.Write "	.frm1.txtBDG_NM.value	 = """ & txtBDG_NM		& """" & vbCr
    Response.Write "	.lgPageNo				 = """ & iIntQueryCount	& """" & vbCr
    Response.Write "	.lgStrPrevKey			 = """ & iStrPrevKey	& """" & vbCr
    Response.Write "	.DbQueryOk	" & vbCr
    Response.Write "End With		" & vbCr
    Response.Write "</Script>		" & vbCr
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : save Data from Db
'============================================================================================================
Sub SubBizSaveMulti()

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear  

Dim PAFG205CUD
Dim Indx,Indx1
Dim arrRowVal,arrColVal
Const C_select_char_sp   = 0 
Const C_acct_ctrl_fg_sp  = 2
Const C_bdg_cd_sp		 = 3    
Const C_gp_acct_nm_sp    = 4
Const C_acct_cd_sp		 = 5
Const C_gp_cd_sp		 = 6
Const C_bdg_ctrl_unit_sp = 7
Const C_trans_fg_sp		 = 8
Const C_divert_fg_sp	 = 9
Const C_add_fg_sp		 = 10
'Const C_bdg_check_fg_sp  = 15

Dim Test
Dim iErrorPosition

Dim IG1_import_group
Const C_select_char = 0 
Const C_bdg_cd = 1    
Const C_acct_cd = 2
Const C_gp_cd = 3
Const C_acct_ctrl_fg = 4
Const C_trans_fg = 5
Const C_divert_fg = 6
Const C_add_fg = 7
Const C_bdg_ctrl_unit = 8
Const C_gp_acct_nm = 9
Const C_bdg_check_fg = 10  
    
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	
	For Indx = 0 To Ubound(arrRowVal) - 1 Step 1
	
	Next

	Redim IG1_import_group(Indx-1,C_bdg_check_fg)
	
	
	
	For Indx1 = 0 To Ubound(arrRowVal) - 1
		arrColVal = Split(arrRowVal(Indx1),gColSep)
		If arrColVal(C_select_char_sp) = "D" Then
			IG1_import_group(Indx1,C_select_char)	= arrColVal(0)
			IG1_import_group(Indx1,C_bdg_cd)		= arrColVal(2)
		Else
			IG1_import_group(Indx1,C_select_char)	= arrColVal(C_select_char_sp)
			IG1_import_group(Indx1,C_bdg_cd)		= arrColVal(C_bdg_cd_sp)
			If arrColVal(C_acct_ctrl_fg_sp) = "A" Then
				IG1_import_group(Indx1,C_acct_cd)		= arrColVal(C_acct_cd_sp)
				IG1_import_group(Indx1,C_gp_cd)			= ""
			ElseIf arrColVal(C_acct_ctrl_fg_sp) = "G" Then			
				IG1_import_group(Indx1,C_acct_cd)		= ""
				IG1_import_group(Indx1,C_gp_cd)			= arrColVal(C_gp_cd_sp)
			End If 			
			IG1_import_group(Indx1,C_acct_ctrl_fg)	= arrColVal(C_acct_ctrl_fg_sp)
			IG1_import_group(Indx1,C_trans_fg)		= arrColVal(C_trans_fg_sp)
			IG1_import_group(Indx1,C_divert_fg)		= arrColVal(C_divert_fg_sp)
			IG1_import_group(Indx1,C_add_fg)		= arrColVal(C_add_fg_sp)
			IG1_import_group(Indx1,C_bdg_ctrl_unit)	= arrColVal(C_bdg_ctrl_unit_sp)
			IG1_import_group(Indx1,C_gp_acct_nm)	= arrColVal(C_gp_acct_nm_sp)
			IG1_import_group(Indx1,C_bdg_check_fg)	= ""	
		End If
	Next

	
    Set PAFG205CUD = Server.CreateObject("PAFG205.cFMngBdgAcctSvr")    
    
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    
    
    Call PAFG205CUD.F_MANAGE_BDG_ACCT_SVR(gStrGlobalCollection,IG1_import_group)
    
	If CheckSYSTEMError(Err, True) = True Then					
       Set PAFG205CUD = Nothing
       Exit Sub
    End If   			
    
    	
    Set PAFG205CUD = nothing    
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
End Sub
%>