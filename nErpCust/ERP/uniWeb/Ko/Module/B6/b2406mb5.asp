<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(부서정보등록)
'*  3. Program ID           : B2406ma1
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Complus List         : 
'                             
'*  7. Modified date(First) : 2005/10/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    On Error Resume Next																	'☜: Protect system from crashing
    Err.Clear																				'☜: Clear Error status

    Call LoadBasisGlobalInf()
    Call HideStatusWnd             

    Dim PB6G061
    Dim I1_after_par_dept_info
    Dim I2_before_par_dept_info
    Dim I3_move_dept_info
    Dim I4_cur_org_id
    
    Const B462_I1_after_par_dept_cd = 0
    Const B462_I1_after_par_dept_lvl = 1
    Const B462_I1_after_par_dept_seq = 2
    
    Const B462_I2_before_par_dept_cd = 0
    Const B462_I2_before_par_dept_lvl = 1
    Const B462_I2_before_par_dept_seq = 2
    
    Const B462_I3_move_dept_cd = 0
    Const B462_I3_move_dept_lvl = 1
    Const B462_I3_move_dept_seq = 2
    
	Redim I1_after_par_dept_info(2)
	Redim I2_before_par_dept_info(2)
	ReDim I3_move_dept_info(2)
	
	I1_after_par_dept_info(B462_I1_after_par_dept_cd)    = Request("txtToParentDeptCd")
	I1_after_par_dept_info(B462_I1_after_par_dept_lvl)	 = Request("txtToParentDeptLvl")
	I1_after_par_dept_info(B462_I1_after_par_dept_seq)	 = Request("txtToParentDeptSeq")

	I2_before_par_dept_info(B462_I2_before_par_dept_cd)  = Request("txtParentDeptCd")
	I2_before_par_dept_info(B462_I2_before_par_dept_lvl) = Request("txtParentDeptLvl")
	I2_before_par_dept_info(B462_I2_before_par_dept_seq) = Request("txtParentDeptSeq")
	
	I3_move_dept_info(B462_I3_move_dept_cd)				 = Request("txtMoveDeptCd")
	I3_move_dept_info(B462_I3_move_dept_lvl)			 = Request("txtMoveDeptLvl")
	I3_move_dept_info(B462_I3_move_dept_seq)			 = Request("txtMoveDeptSeq")	

	I4_cur_org_id										 = Request("txtOrgId")

	Set PB6G061 = server.CreateObject("PB6G061.cBControlHorgMas")

    If CheckSYSTEMError(Err,True) = True Then
        Response.End  
    End If	
    
	Call PB6G061.B_MOVE_DEPT(gStrGlobalCollection,I1_after_par_dept_info,I2_before_par_dept_info,I3_move_dept_info,I4_cur_org_id)
	
    If CheckSYSTEMError(Err,True) = True Then
        Set PB6G061 = Nothing
        Response.End  
    End If	

	Set  PB6G061 = Nothing

%>

<Script Language="VBScript">
    With Parent

       .DBSaveOk
	End With       
</Script>	
