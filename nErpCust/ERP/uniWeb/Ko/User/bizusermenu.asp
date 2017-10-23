<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../inc/IncServer.asp" -->

<Script Language=vbscript src="../inc/incUni2KTV.vbs"></Script>
<%

    Const C_OPMODE         = 0
    Const C_MNU_ID         = 1    ' 0: Menu ID
    Const C_MNU_UPPER      = 2    ' 1: Upper Menu ID
    Const C_MNU_NM         = 3    ' 2: Menu Name
    Const C_MNU_TYPE       = 4    ' 3: Menu Type
    Const C_MNU_LVL        = 5    ' 4: Menu Lvl
    Const C_MNU_SEQ        = 6    ' 5: Menu Seq
    Const C_MNU_PREV_ID    = 7    ' 6: PrevID
    Const C_MNU_PREV_UPPER = 8    ' 7: PrevUpper

    Const C_UNDERBAR       = "_"

	Dim lgOpModeCRUD
	
	On Error Resume Next
	Err.Clear 

	Call HideStatusWnd

	lgOpModeCRUD = Request("txtMode")
	
	Select Case lgOpModeCRUD					
		Case CStr(UID_M0002)
			Call SubBizSaveMulti()	
	End Select 
	
	    

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMulti()
		
	Dim iZC015
	Dim iErrorPosition
	Dim iStrSpread
	
	Const ZC15_I1_MNU_ID       = 0
	Const ZC15_I1_UPPER_MNU_ID = 1
	Const ZC15_I1_MNU_NM       = 2
	Const ZC15_I1_MNU_TYPE     = 3
	Const ZC15_I1_SYS_LVL      = 4
	Const ZC15_I1_MNU_SEQ      = 5
	
	Const ZC15_E1_Mode         = 0	
	Const ZC15_E1_Mnu_Id       = 1
	Const ZC15_E1_Upper_Mnu_Id = 2
	Const ZC15_E1_Mnu_Nm       = 3
	
	'TREE khy20030117
	Const ZC15_E2_Mode = 0
	Const ZC15_E2_Mnu_Id = 1
	Const ZC15_E2_Upper_mnu_Id = 2
	Const ZC15_E2_Mnu_Nm = 3
	Const ZC15_E2_Mnu_Type = 4
	Const ZC15_E2_Sys_Lvl = 5
	Const ZC15_E2_Mnu_Seq = 6
	
	Dim E1_Z_Usr_Mnu 
	
	On Error Resume Next 
	Err.Clear 
	
	iStrSpread = Request("txtMulti")

	Set iZC015 = Server.CreateObject("PZCG015.cCtrlUsrMnu")
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iZC015 = Nothing		                                               
       Exit Sub
    End If
     
	E1_Z_Usr_Mnu = iZC015.ZC_CTRL_USR_MNU (gStrGlobalCollection,iStrSpread,iErrorPosition)
     
    If CheckSYSTEMError2(Err,True,iErrorPosition & "Row","","","","") = True Then  		    
       Set iZC015 = Nothing	
       Exit Sub		
    End If
    
    
    Set iZC015 = Nothing	



End Sub


%>

