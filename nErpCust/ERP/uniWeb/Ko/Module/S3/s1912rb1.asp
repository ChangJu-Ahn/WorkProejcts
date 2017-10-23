<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
    Dim lgOpModeCRUD
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)		                                                 '¢Ð: Query            
             Call SubBizQueryMulti()       
    End Select


'============================================================================================================
Sub SubBizQueryMulti()
	Dim iS19126
	Dim StrNextKey1
	Dim StrNextKey2
	Dim iLngRow
	Dim istrData	
	Dim iLngMaxRow														' S/O Referenc Á¶È¸¿ë Object
    Const C_SHEETMAXROWS_D  = 100
    
    Dim I1_s_available_stock
    Dim I2_b_item
    Dim I3_b_plant 

    Dim E1_b_plant 
    Dim E2_b_item 
    Dim E3_s_available_stock
    Dim EG1_exp_grp 

    Const S335_I1_available_dt = 0              '[CONVERSION INFORMATION]  View Name : imp_next s_available_stock
    Const S335_I1_unique_id = 1 
   
    Const S335_E1_plant_cd               = 0    '[CONVERSION INFORMATION]  View Name : exp b_plant
    Const S335_E1_plant_nm               = 1

    Const S335_E2_basic_unit             = 0    '[CONVERSION INFORMATION]  View Name : exp b_item
    Const S335_E2_item_cd                = 1
    Const S335_E2_item_nm                = 2

'    Const S335_E3_available_dt           = 0    '[CONVERSION INFORMATION]  View Name : exp_next s_available_stock
'    Const S335_E3_unique_id              = 1

    Const S335_EG1_E1_available_dt       = 0    '[CONVERSION INFORMATION]  View Name : exp_item s_available_stock
    Const S335_EG1_E1_on_hand_stock      = 1
    Const S335_EG1_E1_schd_rcpt_qty      = 2
    Const S335_EG1_E1_schd_issue_qty     = 3
    Const S335_EG1_E1_available_stk_qty  = 4

    On Error Resume Next                                                                 '¢Ð: Protect system from crashing
    Err.Clear																			 '¢Ð: Clear Error status                                                            
         
	'-----------------------
	'Data manipulate  area(import view match)
	'-----------------------

	I2_b_item   = Trim(Request("txtItem"))
	I3_b_plant  = Trim(Request("txtPlant"))
	Redim I1_s_available_stock(S335_I1_unique_id)
	I1_s_available_stock(S335_I1_available_dt) = UNIConvDate(Request("lgStrPrevKey1"))
	I1_s_available_stock(S335_I1_unique_id) = Trim(Request("lgStrPrevKey2"))

    Set iS19126 = Server.CreateObject("PS3G119.cSSimulateStkSvr")	

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
 
    call iS19126.S_SIMULATE_ABLE_STK_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D ,I1_s_available_stock,I2_b_item , _
                  I3_b_plant , "" , E1_b_plant , E2_b_item , E3_s_available_stock,  EG1_exp_grp )
	If CheckSYSTEMError(Err,True) = True Then
       Set iS19126 = Nothing		                                                 '¢Ð: Unload Comproxy DLL
       Exit Sub
    End If   
    
    Set iS19126 = Nothing	
    
	istrData = ""
    iLngMaxRow  = CLng(Request("txtMaxRows"))										 '¢Ð: Fetechd Count      
	For iLngRow = 0 To UBound(EG1_exp_grp,1) 
		If  iLngRow < C_SHEETMAXROWS_D  Then

		Else
			StrNextKey1 = EG1_exp_grp(S335_EG1_E1_available_dt)
			StrNextKey2 = E3_s_available_stock
            Exit For
        End If 
        
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_exp_grp(iLngRow,S335_EG1_E1_available_dt))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S335_EG1_E1_available_stk_qty), ggQty.DecPoint, 0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S335_EG1_E1_schd_rcpt_qty), ggQty.DecPoint, 0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S335_EG1_E1_schd_issue_qty), ggQty.DecPoint, 0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_exp_grp(iLngRow,S335_EG1_E1_on_hand_stock), ggQty.DecPoint, 0)     
'        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow 
        istrData = istrData & Chr(11) & Chr(12)                 
	    

	    %>
	    <br>
	    <%
    Next 

'   Response.Write  "A" & iLngRow
'    Response.End 
    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write "With Parent	           " & vbCr
    Response.Write ".txtStockType.value = """ & ConvSPChars(E2_b_item(S335_E2_basic_unit)) & """" & vbCr
 	Response.Write ".ggoSpread.Source = .vspdData           " & vbCr  
    Response.Write ".ggoSpread.SSShowDataByClip   """ & istrData  & """" & vbCr
		
	Response.Write ".lgStrPrevKey1    = """ & StrNextKey1   & """" & vbCr
	Response.Write ".lgStrPrevKey2    = """ & StrNextKey2   & """" & vbCr

	Response.Write ".txtHItem.value   = """ & ConvSPChars(Request("txtItem")) & """" & vbCr
	Response.Write ".txtHPlant.value  = """ & ConvSPChars(Request("txtPlant")) & """" & vbCr
	Response.Write ".DbQueryOk                     " & vbCr  
	Response.Write ".vspdData.focus                " & vbCr  
    Response.Write "End With           " & vbCr															    	
    Response.Write "</Script>          " & vbCr      
     
           	
End Sub    

'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>
