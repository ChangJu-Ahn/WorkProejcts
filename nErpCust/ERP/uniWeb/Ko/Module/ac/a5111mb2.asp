<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<%

On Error Resume Next                                                            '☜: Protect system from crashing
Err.Clear   

Call LoadBasisGlobalInf()   

Dim txtBCtrlCD
Dim txtCCtrlCD
                                                                           '☜: Clear Error status
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim PACG055LKUP

Dim E1_a_ctrl_item
Const C1_TBL_ID = 0
Const C1_COL_ID = 1
Const C1_COL_NM = 2

Dim E2_a_ctrl_item
Const C2_TBL_ID = 0
Const C2_COL_ID = 1
Const C2_COL_NM = 2
	
	Set PACG055LKUP = Server.CreateObject("PACG055.cALkUpCtlItmSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
    End If
    
    txtBCtrlCD = Trim(Request("hBCtrlCd"))
    txtCCtrlCD = Trim(Request("hCCtrlCd"))
    
	Call PACG055LKUP.A_LOOKUP_CTRL_ITEM_SVR(gStrGlobalCollection, txtBCtrlCD, txtCCtrlCD, E1_a_ctrl_item,	E2_a_ctrl_item)
	
	
    If CheckSYSTEMError(Err,True) = True Then
		Set PACG055LKUP = nothing		
		Response.End
    End If
   
	Set PACG055LKUP = nothing 
	
	If isempty(E1_a_ctrl_item) = False And isempty(E2_a_ctrl_item) = False Then
		Response.Write "<Script Language=vbscript>  " & vbCr
	   	Response.Write " with parent.frm1" & vbCr
	   	Response.Write " .hBTblId.value		= """ & ConvSPChars(E1_a_ctrl_item(C1_TBL_ID)) & """										" & vbCr
	   	Response.Write " .hBColmId.value	= """ & ConvSPChars(E1_a_ctrl_item(C1_COL_ID)) & """										" & vbCr
	   	Response.Write " .hBColmIdNm.value	= """ & ConvSPChars(E1_a_ctrl_item(C1_COL_NM)) & """										" & vbCr
	   	Response.Write " .hCTblId.value		= """ & ConvSPChars(E2_a_ctrl_item(C2_TBL_ID)) & """										" & vbCr
	   	Response.Write " .hCColmId.value	= """ & ConvSPChars(E2_a_ctrl_item(C2_COL_ID)) & """										" & vbCr
	   	Response.Write " .hCColmIdNm.Value	= """ & ConvSPChars(E2_a_ctrl_item(C2_COL_NM)) & """										" & vbCr
		Response.Write "End with				" & vbcr
	    Response.Write "Parent.DbPopUpQueryOK		" & vbcr
	    Response.Write "</Script>               " & vbCr
	End If	
	
%>
