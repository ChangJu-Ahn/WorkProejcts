<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : a5401ma1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2006/04/03
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜:
ERR.Clear

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim lgStrPrevKey									' 이전 값 
Dim LngMaxRow										' 현재 그리드의 최대Row
Dim LngRow
Dim lgMaxCount

Call LoadBasisGlobalInf()

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 

'Multi SpreadSheet
LngMaxRow = Request("txtMaxRows")					'☜: Read Operation Mode (CRUD)

Select Case strMode
    Case CStr(UID_M0001)							'☜: Query
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                            '☜: Save,Update             
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)                            '☜: Save,Update             
         Call SubBizCopy()         
End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	Dim iPABG050
	Dim I1_ver_cd
	Dim EG1_exchange_version_info
	Dim iStrData				'조회데이타 저장변수 
'    Dim iIntLoopCount
    Dim iLngRow

'    Dim iStrPrevKey
    
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
'	Const C_SHEETMAXROWS_D = 100		 											'한 화면에 보여지는 최대갯수*1.5

	'###########################################
	'EXPORT GROUP PARAMETER
	Const C_EXP_GRP_EXCH_CD			  = 0
	Const C_EXP_GRP_EXCH_NM		      = 1
	Const C_EXP_GRP_EXCH_USP	      = 2
	Const C_EXP_GRP_EXCH_ACCT_CD	  = 3
	Const C_EXP_GRP_EXCH_ACCT_NM	  = 4
	Const C_EXP_GRP_MASTER_REFLECT_FG = 5
	Const C_EXP_GRP_USP_FG		      = 6	
	'###########################################
        
'    lgMaxCount  = C_SHEETMAXROWS_D						'☜: Fetch count at a time for VspdData
    
    I1_ver_cd = UCASE(Trim(Request("txtVerCd")))

'    lgStrPrevKey = Request("lgStrPrevKey")					'☜: Next Key Value
    
     '###########################################
 
    Set iPABG050 = Server.CreateObject("PABG050.cALkupExchangeVerSvr")

	If CheckSYSTEMError(Err, True) = True Then
		Exit Sub
    End If   
    
	Call iPABG050.A_LIST_EXCHANGE_VERSION_SVR(gStrGlobalCollection, _
	                                      I1_ver_cd, _
	                                      EG1_exchange_version_info)

	If CheckSYSTEMError(Err, True) = True Then					
		Set iPABG050 = Nothing
		Response.Write " <Script Language=vbscript>      " & vbCr
		Response.Write " parent.DbQueryOk                " & vbCr	
		Response.Write "</Script>                        " & vbCr
		Exit Sub
    End If

    Set iPABG050 = Nothing

    '###################################################################################################
   
	iStrData = ""
'    iIntLoopCount = 0	
	
	'###################################################################################################
	For iLngRow = 0 To UBound(EG1_exchange_version_info, 1)
'		iIntLoopCount = iIntLoopCount + 1
'		If  iLngRow < lgMaxCount Then
			istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_EXCH_CD)))
			istrData = istrData & Chr(11) & ""			
			istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_EXCH_NM)))	
			istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_EXCH_USP)))
			istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_EXCH_ACCT_CD)))	
			istrData = istrData & Chr(11) & ""			
			istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_EXCH_ACCT_NM)))
			istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_MASTER_REFLECT_FG)))
			istrData = istrData & Chr(11) & ConvSPChars(Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_USP_FG)))
			istrData = istrData & Chr(11) & Cint(LngMaxRow) + iLngRow + 1
			istrData = istrData & Chr(11) & Chr(12)
'		Else
'			iStrPrevKey = Left(Trim(EG1_exchange_version_info(ILNGROW, C_EXP_GRP_YYYYMM)),4) & Right(Trim(EG1_exchange_version_info(ILNGROW, C_EXP_GRP_YYYYMM)),2) _ 
'			             & Trim(EG1_exchange_version_info(ILNGROW, C_EXP_GRP_COST_TYPE)) & Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_COST_CD)) _
'			             & Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_ACCT_GP)) & Trim(EG1_exchange_version_info(iLngRow, C_EXP_GRP_ACCT_CD))
'		End If		
	Next

'	If iLngRow < lgMaxCount Then
'		iStrPrevKey = ""		
'	End If	

	'#####################################################################################################
	'☜: 화면 처리 ASP 를 지칭함 
	Response.Write " <Script Language=vbscript>				      " & vbCr
    Response.Write " With parent							      " & vbCr
	Response.Write " .frm1.txtVerCd.value =	"""	& I1_ver_cd &  """" & vbCr
	Response.Write " .frm1.htxtVerCd.value = """ & I1_ver_cd & """" & vbCr	
	
	Response.Write " .frm1.vspdData.Redraw = False			      " & vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData	   	      " & vbCr
	Response.Write " .ggoSpread.SSShowData 	""" & istrData &   """" & vbCr	
	Response.Write " .frm1.vspdData.Redraw = True			      " & vbCr	
	Response.Write " .DbQueryOk  							   	  " & vbCr	
	Response.Write " End With								      " & vbCr	
	Response.Write " </Script>								      " & vbCr
'#######################################################################################################


End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	Dim iPABG050
	Dim I1_ver_cd
	Dim iErrorPosition

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Set iPABG050 = Server.CreateObject("PABG050.cAMngExchangeVerSvr")

    If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
    End If    

	I1_ver_cd = Trim(Request("txtVerCd"))

    Call iPABG050.A_MANAGE_EXCHANGE_VERSION_SVR(gStrGlobalCollection, _
											I1_ver_cd, _
											Trim(Request("txtSpread")), _
											iErrorPosition)
		
    If CheckSYSTEMError2(Err, True,iErrorPosition & "행","","","","") = True Then
		Set iPABG050 = Nothing
		Exit Sub
    End If    
    
    Set iPABG050 = Nothing
		
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub	

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizCopy()
	
	Dim iPABG050
	Dim I1_old_ver_cd
	Dim I2_new_ver_cd

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Set iPABG050 = Server.CreateObject("PABG050.cACopyExchangeVerSvr")

    If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
    End If    

	I1_old_ver_cd = Trim(Request("txtVerCd"))
	I2_new_ver_cd = Trim(Request("txtNewVerCd"))

    Call iPABG050.A_COPY_EXCHANGE_VER_SVR(gStrGlobalCollection, _
											I1_old_ver_cd, _
											I2_new_ver_cd)
		
    If CheckSYSTEMError(Err, True) = True Then
		Set iPABG050 = Nothing
		Exit Sub
    End If

    Set iPABG050 = Nothing
		
	Response.Write " <Script Language=vbscript>                       " & vbCr
    Response.Write " With parent				    	              " & vbCr	
	Response.Write " .frm1.txtVerCd.value =	 """ & I2_new_ver_cd & """" & vbCr
	Response.Write " .frm1.htxtVerCd.value = """ & I2_new_ver_cd & """" & vbCr
	Response.Write " .frm1.txtNewVerCd.value = """"" & vbCr				
	Response.Write " .DbSaveOk                                        " & vbCr
    Response.Write " End With                			              " & vbCr	
    Response.Write "</Script>                                         " & vbCr

End Sub	


%>
