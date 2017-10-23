<%@ LANGUAGE="VBSCRIPT" %>
<% Option explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1212MA2
'*  4. Program Name         : 공급처칼렌다수정 
'*  5. Program Desc         : 공급처칼렌다수정 
'*  6. Component List       : PM1G228.cMListSpplCalS / PM1G221.cMMaintSpplCalS
'*  7. Modified date(First) : 2001/01/16
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Min, Hak-jun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->

<%
    Dim lgOpModeCRUD
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	Call LoadBasisGlobalInf()

    lgOpModeCRUD  = Request("txtMode") 
										                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
    Dim iM12128

	Dim idtDate
	Dim istartIndex
	Dim lastDay
	Dim i
	Dim CalCol

	Dim I1_b_biz_partner
	Dim I2_m_str_wks
	Dim E1_b_biz_partner
	Dim EG1_exp_group
	
	Const M307_EG1_E1_m_sppl_cal_cal_dt = 0
    Const M307_EG1_E1_m_sppl_cal_holi_type = 1
    Const M307_EG1_E1_m_sppl_cal_day_of_week = 2
    Const M307_EG1_E1_m_sppl_cal_holi_desc = 3
    ReDim EG1_exp_group(31, M307_EG1_E1_m_sppl_cal_holi_desc)

    Const M307_E1_bp_cd = 0
    Const M307_E1_bp_nm = 1
    ReDim E1_b_biz_partner(M307_E1_bp_nm)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	If Request("txtYear") = "" Or Request("txtMonth") = "" Then				'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("229903", vbInformation, "", "", I_MKSCRIPT)          
		Response.End 
	End If
	
    Set iM12128 = Server.CreateObject("PM1G228.cMListSpplCalS")
	
	
			If CheckSYSTEMError(Err,True) = true Then 		
			 	Set iM12128 = Nothing												'☜: ComProxy Unload
			Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
			End if
    
    Dim strYYYYMM
    strYYYYMM = Right("0000" & Request("txtYear"), 4)
    strYYYYMM = strYYYYMM & "-"
    strYYYYMM = strYYYYMM & Right("00" & Request("txtMonth"), 2)
	
	I1_b_biz_partner = Trim(Request("txtBpCd"))
	I2_m_str_wks = strYYYYMM  


    call iM12128.M_LIST_SPPL_CAL_SVR(gStrGlobalCollection, I1_b_biz_partner, I2_m_str_wks, E1_b_biz_partner, EG1_exp_group)

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.frm1.txtBpNm.value = """ & ConvSPChars(E1_b_biz_partner(M307_E1_bp_nm))      & """" & vbCr
    Response.Write "</Script>"                  & vbCr 

    
	If CheckSYSTEMError(Err, True) = True then	
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " Dim CalCol " & vbCr
	Response.Write " For CalCol = 0 to 41 "     & vbCr
	Response.Write "Parent.frm1.txtDate(CalCol).value = "" ""      "     & vbCr 
	Response.Write "Parent.frm1.txtDate(CalCol).className = ""DummyDay""" & vbCr
	Response.Write "Parent.frm1.txtDate(CalCol).disabled = True "        & vbCr
	
	Response.Write "Parent.frm1.txtHoli(CalCol).value = "" """           & vbCr
	Response.Write "Parent.frm1.txtHoli(CalCol).disabled = True "        & vbCr
	
	Response.Write "Parent.frm1.txtDesc(CalCol).value = "" """           & vbCr
	Response.Write "Parent.frm1.txtDesc(CalCol).disabled = True "        & vbCr
	Response.Write "Parent.frm1.txtDesc(CalCol).title = "" """           & vbCr
	Response.Write "Next"                       & vbCr
    Response.Write "</Script>"                  & vbCr    

		Set iM12128 = Nothing																	'☜: ComProxy UnLoad
		Exit Sub																		'☜: Process End
	End If
	'-----------------------
	'Result data display area
	'----------------------- 
    idtDate = CDate(Left(EG1_exp_group(0, M307_EG1_E1_m_sppl_cal_cal_dt), 10))
    istartIndex = EG1_exp_group(0, M307_EG1_E1_m_sppl_cal_day_of_week)-1
    
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.frm1.hYear.value = """ & Year(idtDate) & """" & vbCr
	Response.Write "Parent.frm1.hMonth.value = """ & Month(idtDate) & """" & vbCr
	Response.Write "Parent.lgStartIndex = """ & istartIndex & """" & vbCr
	Response.Write "</Script>"                  & vbCr    

	idtDate = DateAdd("d", -1, idtDate)
	lastDay = Day(idtDate)

	'지난달 Display를 위해서....
    idtDate = CDate(EG1_exp_group(0, M307_EG1_E1_m_sppl_cal_cal_dt))
	
	'1일 이전 데이타 클리어---
    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "Parent.lgLastDay = """ & lastDay & """" & vbCr
    Response.Write "Dim CalCol " & vbCr
    Response.Write "For CalCol = " & istartIndex-1 & " To 0 Step -1 " & vbCr
  
    Response.Write "Parent.frm1.txtDate(CalCol).value = CStr(" & lastDay & " + CalCol - " & istartIndex-1 & ")" & vbCr 
	Response.Write "Parent.frm1.txtDate(CalCol).className = ""DummyDay""" & vbCr
	Response.Write "Parent.frm1.txtDate(CalCol).disabled = True "         & vbCr
	
	Response.Write "Parent.frm1.txtHoli(CalCol).value = "" """            & vbCr
	Response.Write "Parent.frm1.txtHoli(CalCol).disabled = True "         & vbCr
	
	Response.Write "Parent.frm1.txtDesc(CalCol).value = "" """            & vbCr
	Response.Write "Parent.frm1.txtDesc(CalCol).disabled = True "         & vbCr
	Response.Write "Parent.frm1.txtDesc(CalCol).title = "" """            & vbCr
	Response.Write "Next"                       & vbCr
	For i = 0 To UBound(EG1_exp_group, 1)
		If EG1_exp_group(i, M307_EG1_E1_m_sppl_cal_holi_type) = "H" Then
	Response.Write "Parent.frm1.txtDate(" & istartIndex & ").style.color = ""red""" & vbCr

		Else
			If (istartIndex + 1) Mod 7 = 0 Then
	Response.Write "Parent.frm1.txtDate(" & istartIndex & ").style.color = ""blue""" & vbCr
			Else
	Response.Write "Parent.frm1.txtDate(" & istartIndex & ").style.color = ""black""" & vbCr
			End If
		End If
	Response.Write "Parent.frm1.txtDate(" & istartIndex & ").value = """ & (i+1) & """" & vbCr
	Response.Write "Parent.frm1.txtDate(" & istartIndex & ").className = ""Day""" & vbCr
	Response.Write "Parent.frm1.txtDate(" & istartIndex & ").disabled = False " & vbCr
	
	Response.Write "Parent.frm1.txtHoli(" & istartIndex & ").value = """ & ConvSPChars(EG1_exp_group(i, M307_EG1_E1_m_sppl_cal_holi_type)) & """" & vbCr
	Response.Write "Parent.frm1.txtHoli(" & istartIndex & ").disabled = False " & vbCr
	
	Response.Write "Parent.frm1.txtDesc(" & istartIndex & ").value = """ & ConvSPChars(EG1_exp_group(i, M307_EG1_E1_m_sppl_cal_holi_desc)) & """" & vbCr
	Response.Write "Parent.frm1.txtDesc(" & istartIndex & ").disabled = False " & vbCr
	Response.Write "Parent.frm1.txtDesc(" & istartIndex & ").title = """ & ConvSPChars(EG1_exp_group(i, M307_EG1_E1_m_sppl_cal_holi_desc)) & """" & vbCr
		istartIndex = istartIndex + 1
	Next
	
	Response.Write " For CalCol = " & istartIndex & " To 41 " & vbCr
	Response.Write "Parent.frm1.txtDate(CalCol).value = CStr(CalCol - " & istartIndex-1 & ")" & vbCr 
	Response.Write "Parent.frm1.txtDate(CalCol).className = ""DummyDay""" & vbCr
	Response.Write "Parent.frm1.txtDate(CalCol).disabled = True "         & vbCr
	
	Response.Write "Parent.frm1.txtHoli(CalCol).value = "" """            & vbCr
	Response.Write "Parent.frm1.txtHoli(CalCol).disabled = True "         & vbCr
	
	Response.Write "Parent.frm1.txtDesc(CalCol).value = "" """           & vbCr
	Response.Write "Parent.frm1.txtDesc(CalCol).disabled = True "        & vbCr
	Response.Write "Parent.frm1.txtDesc(CalCol).title = "" """           & vbCr
	Response.Write "Next"                        & vbCr
	
	Response.Write "Parent.lgNextNo = """""      & vbCr   
	Response.Write "Parent.lgPrevNo = """""      & vbCr
	Response.Write "Parent.DbQueryOk "           & vbCr
	
    Response.Write "</Script>"                  & vbCr
	
    Set iM12128 = Nothing															'☜: Unload Comproxy
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim iM12221															'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

	Dim idtDate
	Dim i
	Dim I1_b_biz_partner
	Dim IG1_imp_group

	Const M306_IG1_I1_m_sppl_cal_cal_dt = 0
	Const M306_IG1_I1_m_sppl_cal_holi_type = 1
	Const M306_IG1_I1_m_sppl_cal_holi_desc = 2
	ReDim IG1_imp_group(31, M306_IG1_I1_m_sppl_cal_holi_desc)
	
	On Error Resume Next  				
    Err.Clear																		'☜: Protect system from crashing

    If Request("txtFlgMode") = "" Then											'⊙: 저장을 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("lgIntFlgMode 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Exit Sub
	End If
	
    Set iM12221 = Server.CreateObject("PM1G221.cMMaintSpplCalS")    
    
    '-----------------------
    'Data manipulate area
    '----------------------- 
    For i = 1 To Request("txtHoli").count
		idtDate = CDate(CStr(Request("hYear") & "-" & Request("hMonth") & "-" & i))
		'idtDate = left(idtDate, 10)	
		I1_b_biz_partner				= Trim(Request("txtBpCd"))
        IG1_imp_group(i, M306_IG1_I1_m_sppl_cal_cal_dt)			= idtDate
        IG1_imp_group(i, M306_IG1_I1_m_sppl_cal_holi_type)		= Request("txtHoli")(i)
        IG1_imp_group(i, M306_IG1_I1_m_sppl_cal_holi_desc)		= Request("txtDesc")(i)
   
    Next

	'-----------------------
	'Com Action Area
	'-----------------------
	Call iM12221.M_MAINT_SPPL_CAL_SVR(gStrGlobalCollection, I1_b_biz_partner, IG1_imp_group)

    If Err.Number <> 0 Then
		Set iM12221 = Nothing																'☜: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙:
		Exit Sub																		'☜: Process End
	End If
		
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	
    if CheckSYSTEMError(Err,True) = true then 
	
		Set iM12221 = Nothing
		Exit Sub
	End If

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.DBSaveOK "           & vbCr
    Response.Write "</Script>"                  & vbCr 
				
    Set iM12221 = Nothing                                                   '☜: Unload Comproxy
        
End Sub    

%>
