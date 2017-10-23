<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	lgErrorStatus  = ""	'>> air


	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")  
	Call HideStatusWnd


    Select Case Request("txtMode")
        Case CStr(UID_M0001)                                                         '��: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMulti()

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
	
	Dim lgErrorPos
    Dim PAFG415CUD
    Dim I2_f_ln_info
    Dim importArray
    Dim idx
	Dim arrColVal
    
    lgErrorPos        = ""   'LN000003
    
    
	If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbOKOnly, "", "", I_MKSCRIPT)		'txtMaxRows ���ǰ��� ����ֽ��ϴ�!
		Response.End
	End If

    Set PAFG415CUD = Server.CreateObject("PAFG415.cFMngLnPlnAnRsltSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
    
    importArray = Split(Trim(Request("txtSpread")), gRowSep)

    Call PAFG415CUD.F_MANAGE_LN_PLAN_AND_RESULT_SVR(gStrGloBalCollection, importArray, I2_f_ln_info)		

    If CheckSYSTEMError(Err, True) = True Then
       Set PAFG415CUD = Nothing
       Exit Sub
    End If    

    Set PAFG415CUD = Nothing



'	| NOTE***********************************************************************
'	| - ������ Biz Logic�� �������� �ʰ�, "���޿�����(�ڱ�)"�� ���尡���ϰ� ����.
'	| - 20080307 >> air
'	| NOTE***********************************************************************
	Call SubOpenDB(lgObjConn)  
	
    For iDx = 0 To UBound(importArray,1) - 1
        
        arrColVal = Split(importArray(iDx), gColSep)                       '��: Split Column data
   'Call ServerMesgBox(arrColVal(0) , vbInformation, I_MKSCRIPT)       
        Select Case arrColVal(0)
            Case "C", "U"  
            	Call SubBizSaveMultiUpdate(arrColVal)                      '��: Update
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next
	'***********************************************************************
    Call SubCloseDB(lgObjConn)  
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
	
End Sub


'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    Dim lgStrSQL
    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
	'dim ii


    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    	
	'for ii = 0 to Ubound(arrColVal,1)
		
		lgStrSQL = ""
		lgStrSQL = lgStrSQL & " UPDATE	F_LN_PLAN SET " 
		lgStrSQL = lgStrSQL & " 				PLAN_LOC_AMT = " & UNIConvNum(arrColVal(6),0)
		lgStrSQL = lgStrSQL & " 		WHERE	LOAN_NO      = " &	FilterVar(Trim(UCase(arrColVal(13))),"''","S")	
		lgStrSQL = lgStrSQL & " 		AND		PAY_PLAN_DT  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")	
		lgStrSQL = lgStrSQL & " 		AND		PAY_OBJ      = " &  FilterVar(Trim(UCase(arrColVal(4))),"''","S")	
		'
'Call ServerMesgBox(lgStrSQL , vbInformation, I_MKSCRIPT)
		
	'Next
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If

End Sub
%>
