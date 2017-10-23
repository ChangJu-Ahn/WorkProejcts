<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>


<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<% 
	Call LoadBasisGlobalInf() 

    On Error Resume Next
    Dim strMode
    DIM FileName

    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
    strMode      = Trim(Request("txtMode"))                                          '��: Read Operation Mode (CRUD)
 

 Dim strFilePath
 
    Select Case strMode
        Case CStr(UID_M0001)                                                         '��: Query
		
            Call SubMakeDisk()
        Case CStr(UID_M0002)  
			Call SubFileDownLoad()          
          
        Case CStr(UID_M0003)                                                         '��: Delete
            Call SubFileDownLoad2()
          
    End Select
   
 
'-----------------------------------------------------------------------------------------------
'            File DownLoad(With B.A)
'-----------------------------------------------------------------------------------------------
Sub SubFileDownLoad()
		On Error Resume Next
		Err.Clear 

		Call HideStatusWnd

		strFilePath = "http://" & Request.ServerVariables("SERVER_NAME") & ":" _
					   & Request.ServerVariables("SERVER_PORT")
        If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
            strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
        End If
		strFilePath = strFilePath  & "files/u2000/"
		strFilePath = strFilePath & Request("txtFileName")
		
End Sub

'--------------------------------------------------------------------------------------------------------
'					            File DownLoad(With xinSoft solution)
'--------------------------------------------------------------------------------------------------------
Sub SubFileDownLoad2()
		Err.Clear                                                               '��: Protect system from crashing
		On Error Resume Next
		Dim xdn

		Call HideStatusWnd

		set xdn = Server.CreateObject("Xionsoft.XionFileDownLoad")

		'strFilePath = "/" & Mid(Request.ServerVariables("URL"), 2, instr(2, Request.ServerVariables("URL"), "/") - 2) & "/" & glang & "/files/" & gCompany & "/"
		'strFilePath = "/" & Mid(Request.ServerVariables("URL"), 2, instr(2, Request.ServerVariables("URL"), "/") - 2) & "/"
		strFilePath = "/" & Mid(Request.ServerVariables("URL"), 2, instr(2, Request.ServerVariables("URL"), "/") - 2) & "/template/files/" & gCompany & "/"
		
		'Call ServerMesgBox(strFilePath, vbInformation, I_MKSCRIPT)

		'���� �� �� �ӽ� COMMENT 2001/1/27
		xdn.DownFromFile strFilePath & Trim(Request("txtFileName"))
		xdn.OnEndPage

		set xdn = nothing
		'Call ServerMesgBox(183114, vbInformation, I_MKSCRIPT)
		'Call DisplayMsgBox("183114","X","X","X")
		Response.End


End Sub


'--------------------------------------------------------------------------------------------------------
'					         
'--------------------------------------------------------------------------------------------------------

Sub SubMakeDisk()
		On Error Resume Next
		Err.Clear

		Dim IArrData 
		Dim iPAVG010

		Dim ex1_file_name
		Dim ex2_return_code

		Const I1_ief_supplied   = 0
    	Const I1_biz_area_cd	= 1
		Const I1_start_issued_dt =2
		Const I1_end_issued_dt =3
		Const I1_report_issued_dt =4
		Const I1_file_name =5
		Const I1_file_path =6
		Const I1_B_start_issued_dt =7
		Const I1_B_end_issued_dt =8
		Const I1_chkYn =9

        Redim IArrData(I1_chkYn)

		 IArrData(I1_ief_supplied)      =  UCase(Trim(Request("cbofileGubun")))        'A:�Ϲ� ���ϻ���/B:������ ���ϻ��� ������/C:�Ϲ�+������ 

		 IArrData(I1_biz_area_cd)    	=	UCase(Trim(Request("txtBizAreaCD")))
		 IArrData(I1_start_issued_dt)   =	UNIConvDate(Request("txtIssueDt1"))
		 IArrData(I1_end_issued_dt)  	=	UNIConvDate(Request("txtIssueDt2"))

		 IArrData(I1_report_issued_dt)	=	UNIConvDate(Request("txtReportDt"))
		 IArrData(I1_file_name)			=	UCase(Request("txtFileName")) & ""

		 strFilePath = Server.MapPath("../../files/u2000") & "\"

		 IArrData(I1_file_path)			=	 strFilePath

		 IArrData(I1_B_start_issued_dt)	=	 UNIConvDate(Request("txtIssueDt5"))
		 IArrData(I1_B_end_issued_dt)	=	 UNIConvDate(Request("txtIssueDt6"))
		 IArrData(I1_chkYn)	=	 Request("chkYn")


		Set iPAVG010 = Server.CreateObject("PAVG010.cbAVatDiskSvrEab")

		If CheckSYSTEMError(Err,True) = True Then
			Response.End
			Exit Sub
		End If

		Call iPAVG010.EAB_A_VAT_DISK_SVR(gStrGlobalCollection,IArrData,ex1_file_name,ex2_return_code)

		If CheckSYSTEMError(Err, True) = True Then
			Set iPAVG010 = Nothing
			Response.End
			Exit Sub
		End If


		'Select Case  Trim(EX2_return_code)
		'	Case "A"	' �Ű����������� ã�� ��  �����ϴ�.
		'		Call ServerMesgBox("�Ű����������� ã�� �� �����ϴ�." , vbInformation, I_MKSCRIPT)
				'Call DisplayMsgBox("700106","X","X","X","X")
		'	Case "B"	' �����ڷ������� ã�� ��  �����ϴ�.
		'		Call ServerMesgBox("�����ڷ������� ã�� ��  �����ϴ�." , vbInformation, I_MKSCRIPT)
		'		'Call DisplayMsgBox("700107","X","X","X")
		'	Case "C"	' �����ڷ������� ã�� ��  �����ϴ�.
		'		'Call DisplayMsgBox("700108","X","X","X")
		'		Call ServerMesgBox("�����ڷ������� ã�� ��  �����ϴ�." , vbInformation, I_MKSCRIPT)
		'	Case "D"	' �ΰ��������� ó���� �Ϸ���� �ʾҽ��ϴ�. �ٽ� ���� �Ͻʽÿ�.
		'		Call ServerMesgBox("�ΰ��������� ó���� �Ϸ���� �ʾҽ��ϴ�. �ٽ� ���� �Ͻʽÿ�." , vbInformation, I_MKSCRIPT)
		'		'Call DisplayMsgBox("700109","X","X","X")
		'	Case "Z"	' ���ȭ�� �ٿ�ε� �۾� 

				FileName=ex1_file_name


				Call HideStatusWnd
				Set iPAVG010 = Nothing
				Response.Write " <SCRIPT LANGUAGE=VBSCRIPT>" & vbCr

				'Response.write " parent.frm1.txtFileName.value = " &  ex1_file_name & """ & vbCr
				Response.write " parent.frm1.txtFileName.value ="""& ex1_file_name &"""" & vbCr
				Response.write " parent.subVatDiskOK(""" & FileName & """)" & vbCr
				response.write "</SCRIPT>" & vbCr

		'End Select

		Call HideStatusWnd

		Set iPAVG010 = Nothing		'��: Unload Comproxy
		Response.End												

End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>

<script language="vbscript">
	Dim SF
	On Error Resume Next
	Err.Clear
	'parent.frm1.txtFileName.value = "<%=ex1_file_name%>"
	Set SF = CreateObject("uni2kCM.SaveFile")
        If SF.SaveTextFile("<%= strFilePath %>") = True Then
			Set SF = Nothing
			 parent.subVatDiskOK2("OK")
		Else
			Set SF = Nothing
			 'parent.subVatDiskOK2("FAIL")
		End If

	
</script>
