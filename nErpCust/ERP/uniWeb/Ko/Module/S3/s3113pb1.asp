<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ְ��� 
'*  3. Program ID           : S3113PA1
'*  4. Program Name         : CTP Check
'*  5. Program Desc         :
'*  6. Comproxy List        : S31141SoSchdLine, uniAPS
'*  7. Modified date(First) : 2000/09/27
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Cho Song-Hyon
'* 10. Modifier (Last)      : Cho Song-Hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/09/27 : CTP Date
'*                            -2001/12/18 : Date ǥ������ 
'**********************************************************************************************

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
On Error Resume Next														'��: 

Call HideStatusWnd

Dim objCon																	'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'Dim txtServer
'Dim txtPort
'Dim txtUsrID
'Dim txtPwd
Dim strPlantCd
Dim strAPSHost
Dim strAPSPort
Dim strOrder
Dim pS31141

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strOrder = ""

strPlantCd = Request("txtPlantCd")
strAPSHost = Request("txtAPSHost")
strAPSPort = Request("txtAPSPort")

Select Case strMode

Case CStr("CTPQuery")														'��: CTPQuery ��ȸ ��û�� ���� 

	Dim retVal

    Err.Clear                                                               

    If Trim(Request("txtSoNo")) = "" Or Trim(Request("txtSoSeq")) = "" Or Trim(Request("txtItemCode")) = "" Then	'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("��ȸ ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
	'pis test
    'Set objCon = Server.CreateObject("uniAPS.APSConnect")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set objCon = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If

	'-----------------------
    'Connection APS Server
    '-----------------------
	strOrder = strOrder & Trim(Request("txtSoNo")) _
				& "-" & Trim(Request("txtSoSeq")) & gColSep					'Order ID - Distnct
	strOrder = strOrder & "" & gColSep										'Order Description
	strOrder = strOrder & "0" & gColSep										'Order Category
	strOrder = strOrder & UNIConvDate(Request("txtTodayDate")) & gColSep	'Order Entry Date
	strOrder = strOrder & UNIConvDate(Request("txtReqDate")) & gColSep		'Order Request Date
	strOrder = strOrder & "" & gColSep										'Order Promise Date
	strOrder = strOrder & "0" & gColSep										'Order Flag - Customer Order : 0
	strOrder = strOrder & "" & gColSep										'Request("txtCustomer")
	
	strOrder = strOrder & Trim(Request("txtItemCode")) & gColSep			'Part ID
	strOrder = strOrder & UNIConvNum(Request("txtReqQty"),0) & gColSep		'Part Qty
	strOrder = strOrder & Trim(Request("txtTrackingNO")) & gRowSep			'Tracking No

'== Global ����(APSHOST,APSPORT : ����� ������ ���)�� incServer.asp�� ����Ǹ� �Ʒ� �ּ�ó���� ������ ����.. ��..
	
	'pis test
	'retVal = objCon.RunCTP2(strPlantCd,strOrder,gDBServer,gDatabase,gDBLoginID,gDBLoginPwd,strAPSHost,strAPSPort)

'==	retVal = objCon.RunCTP2(strPlantCd,strOrder,gDBServer,gDatabase,gDBLoginID,gDBLoginPwd,APSHOST,APSPORT)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)
	   Set objCon = Nothing																	'��: ComProxy UnLoad
		%>
		<Script Language=vbscript>
			parent.frm1.btnCTPSave.Disabled = True
		</Script>
		<%
	   Response.End																				'��: Process End
	End If

	If objCon.ErrorNo <> 0 Then
		Call ServerMesgBox(objCon.ErrorText, vbCritical, I_MKSCRIPT)              
		Set objCon = Nothing
		%>
		<Script Language=vbscript>
			parent.frm1.btnCTPSave.Disabled = True
		</Script>
		<%
		Response.End 
	End If

	Dim arrTemp
	Dim arrVal
	Dim AccQty
	Dim ReqQty
	Dim ReqDate


	arrTemp = Split(retVal, gRowSep)
	arrVal = Split(arrTemp(0), gColSep)

	AccQty = ""
	AccQty = UNICDbl(arrVal(1),0)
	AccQty = AccQty + UNICDbl(arrVal(2),0)
	AccQty = AccQty + UNICDbl(arrVal(3),0)
	AccQty = UNIConvNum(AccQty,0)
    
	ReqQty = UNICDbl(Request("txtReqQty"),0)
	ReqDate = UNIConvDate(Request("txtReqDate"))

	'-----------------------
	'Result data display area
	'----------------------- 
	
%>
<Script Language=vbscript>
	With parent.frm1

		<% '���������� ���� %>
		.txtAccDate_All.value	= "<%=UNIDateClientFormat(arrVal(0))%>"
		<% '�䱸�Ǵ� �ѷ� %>
		.txtAccQty_All.value	= "<%=UNINumClientFormat(ReqQty,ggQty.DecPoint,0)%>"			

		<% '�䱸�Ǵ� ����(12/05) >= ���������� ����(12/04 or 12/05) %>
		If "<%=UNIDateClientFormat(ReqDate)%>" >= "<%=UNIDateClientFormat(arrVal(0))%>" Then

			<% '���������� ���� %>
			.txtAccDate_Sub1.value	= "<%=UNIDateClientFormat(arrVal(0))%>"
			<% '�䱸�Ǵ³�¥�� ���� ������ ���� %>
			.txtAccQty_Sub1.value	= "<%=UNINumClientFormat(AccQty,ggQty.DecPoint,0)%>"
			<% '�䱸�Ǵ� ���� %>
			.txtAccDate_Sub2.value	= "<%=UNIDateClientFormat(ReqDate)%>"
			<% '�Ѽ��� - ���������� ���� %>
			.txtAccQty_Sub2.value	= 0

			<% '���Ҽ����� Protect %>
			Call parent.ggoOper.SetReqAttr(.rdoSelect_Sub, "Q")

		<% '�䱸�Ǵ� ����(12/05) < ���������� ����(12/06) %>
		ElseIf "<%=UNIDateClientFormat(ReqDate)%>" < "<%=UNIDateClientFormat(arrVal(0))%>" Then

			<% '�䱸�Ǵ� ���� %>
			.txtAccDate_Sub1.value	= "<%=UNIDateClientFormat(ReqDate)%>"							
			<% '�䱸�Ǵ³�¥�� ���� ������ ���� %>
			.txtAccQty_Sub1.value	= "<%=UNINumClientFormat(AccQty,ggQty.DecPoint,0)%>"
			<% '���������� ���� %>
			.txtAccDate_Sub2.value	= "<%=UNIDateClientFormat(arrVal(0))%>"										
			<% '�Ѽ��� - ���������� ���� %>
			.txtAccQty_Sub2.value	= "<%=UNINumClientFormat(ReqQty - AccQty,ggQty.DecPoint,0)%>"
		End If

	End With
</Script>
<%
	 
    'Set objCon = Nothing															'��: Unload Comproxy
	Response.End																	'��: Process End


Case CStr("CTPAccept")																'��: CTP�� ���� ��û 

	Dim ProjDate
								
    Err.Clear																		

    'Set objCon = Server.CreateObject("uniAPS.APSConnect")    
		   
	'-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set objCon = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	'-----------------------
    'Connection APS Server
    '-----------------------
	Select Case Request("txtRadioFlg")
    Case "A"
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & gColSep						'Order ID - Distnct
		strOrder = strOrder & "" & gColSep											'Order Description
		strOrder = strOrder & "0" & gColSep											'Order Category
		strOrder = strOrder & UNIConvDate(Request("txtTodayDate")) & gColSep		'Order Entry Date
		strOrder = strOrder & UNIConvDate(Request("txtReqDate")) & gColSep			'Order Request Date
		strOrder = strOrder & UNIConvDate(Request("txtAccDate_All")) & gColSep		'Order Promise Date
		strOrder = strOrder & "0" & gColSep											'Order Flag - Customer Order : 0
		strOrder = strOrder & "" & gColSep											'Request("txtCustomer")
	
		strOrder = strOrder & Trim(Request("txtItemCode")) & gColSep				'Part ID
		strOrder = strOrder & UNIConvNum(Request("txtAccQty_All"),0) & gColSep		'Part Qty
		strOrder = strOrder & Trim(Request("txtTrackingNO")) & gRowSep				'Tracking No

	Case "S"
		'--First Date AcceptOrder
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & "-1" & gColSep				'Order ID - Distnct
		strOrder = strOrder & "" & gColSep											'Order Description
		strOrder = strOrder & "0" & gColSep											'Order Category
		strOrder = strOrder & UNIConvDate(Request("txtTodayDate")) & gColSep		'Order Entry Date
		strOrder = strOrder & UNIConvDate(Request("txtReqDate")) & gColSep			'Order Request Date
		strOrder = strOrder & UNIConvDate(Request("txtAccDate_Sub1")) & gColSep		'Order Promise Date
		strOrder = strOrder & "0" & gColSep											'Order Flag - Customer Order : 0
		strOrder = strOrder & "" & gColSep											'Request("txtCustomer")
	
		strOrder = strOrder & Trim(Request("txtItemCode")) & gColSep				'Part ID
		strOrder = strOrder & UNIConvNum(Request("txtAccQty_Sub1"),0) & gColSep		'Part Qty
		strOrder = strOrder & Trim(Request("txtTrackingNO")) & gRowSep				'Tracking No

		'--Second Date AcceptOrder
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & "-2" & gColSep				'Order ID - Distnct
		strOrder = strOrder & "" & gColSep											'Order Description
		strOrder = strOrder & "0" & gColSep											'Order Category
		strOrder = strOrder & UNIConvDate(Request("txtTodayDate")) & gColSep		'Order Entry Date
		strOrder = strOrder & UNIConvDate(Request("txtReqDate")) & gColSep			'Order Request Date
		strOrder = strOrder & UNIConvDate(Request("txtAccDate_Sub2")) & gColSep		'Order Promise Date
		strOrder = strOrder & "0" & gColSep											'Order Flag - Customer Order : 0
		strOrder = strOrder & "" & gColSep											'Request("txtCustomer")
	
		strOrder = strOrder & Trim(Request("txtItemCode")) & gColSep				'Part ID
		strOrder = strOrder & UNIConvNum(Request("txtAccQty_Sub2"),0) & gColSep		'Part Qty
		strOrder = strOrder & Trim(Request("txtTrackingNO")) & gRowSep				'Tracking No

	End Select

'== Global ����(APSHOST,APSPORT : ����� ������ ���)�� incServer.asp�� ����Ǹ� �Ʒ� �ּ�ó���� ������ ����.. ��..
	
	'pis test
	'ProjDate = objCon.AcceptOrder(strPlantCd,strOrder,gDBServer,gDatabase,gDBLoginID,gDBLoginPwd,strAPSHost,strAPSPort)

'==	ProjDate = objCon.AcceptOrder(strPlantCd,strOrder,gDBServer,gDatabase,gDBLoginID,gDBLoginPwd,APSHOST,APSPORT)

		   
	If Err.Number <> 0 Then
		Set objCon = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	If objCon.ErrorNo <> 0 Then
		Call ServerMesgBox(objCon.ErrorText, vbCritical, I_MKSCRIPT)              
		Set objCon = Nothing
		Response.End 
	End If


	Dim arrProjVal
	arrProjVal = Split(ProjDate, gRowSep)

%>
<Script Language=vbscript>
	With parent.frm1
		<% '������(OLD)���������� ���� %>
		.txtBeforeChangeDate.value	= .txtAccDate_All.value
		<% '������(NEW)���������� ���� %>
		.txtAfterChangeDate.value	= "<%=UNIDateClientFormat(arrProjVal(0))%>"
		parent.DbCTPSaveOk
	End With
</Script>
<%

	Set objCon = Nothing
	Response.End 	

Case CStr("CTPModify")																'��: CTP�� ���� ��û 

	Dim ProjModDate
								
    Err.Clear																		

    'Set objCon = Server.CreateObject("uniAPS.APSConnect")    
		   
	'-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set objCon = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
	
	'-----------------------
    'Connection APS Server
    '-----------------------
	Select Case Request("txtRadioFlg")
    Case "A"
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & gColSep						'Order ID - Distnct
		strOrder = strOrder & "" & gColSep											'Order Description
		strOrder = strOrder & "0" & gColSep											'Order Category
		strOrder = strOrder & UNIConvDate(Request("txtTodayDate")) & gColSep		'Order Entry Date
		strOrder = strOrder & UNIConvDate(Request("txtReqDate")) & gColSep			'Order Request Date
		strOrder = strOrder & UNIConvDate(Request("txtAccDate_All")) & gColSep		'Order Promise Date
		strOrder = strOrder & "0" & gColSep											'Order Flag - Customer Order : 0
		strOrder = strOrder & "" & gColSep											'Request("txtCustomer")
	
		strOrder = strOrder & Trim(Request("txtItemCode")) & gColSep				'Part ID
		strOrder = strOrder & UNIConvNum(Request("txtAccQty_All"),0) & gColSep		'Part Qty
		strOrder = strOrder & Trim(Request("txtTrackingNO")) & gRowSep				'Tracking No

	Case "S"
		'--First Date ModifyOrder
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & "-1" & gColSep				'Order ID - Distnct
		strOrder = strOrder & "" & gColSep											'Order Description
		strOrder = strOrder & "0" & gColSep											'Order Category
		strOrder = strOrder & UNIConvDate(Request("txtTodayDate")) & gColSep		'Order Entry Date
		strOrder = strOrder & UNIConvDate(Request("txtReqDate")) & gColSep			'Order Request Date
		strOrder = strOrder & UNIConvDate(Request("txtAccDate_Sub1")) & gColSep		'Order Promise Date
		strOrder = strOrder & "0" & gColSep											'Order Flag - Customer Order : 0
		strOrder = strOrder & "" & gColSep											'Request("txtCustomer")
	
		strOrder = strOrder & Trim(Request("txtItemCode")) & gColSep				'Part ID
		strOrder = strOrder & UNIConvNum(Request("txtAccQty_Sub1"),0) & gColSep		'Part Qty
		strOrder = strOrder & Trim(Request("txtTrackingNO")) & gRowSep				'Tracking No

		'--Second Date ModifyOrder
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & "-2" & gColSep				'Order ID - Distnct
		strOrder = strOrder & "" & gColSep											'Order Description
		strOrder = strOrder & "0" & gColSep											'Order Category
		strOrder = strOrder & UNIConvDate(Request("txtTodayDate")) & gColSep		'Order Entry Date
		strOrder = strOrder & UNIConvDate(Request("txtReqDate")) & gColSep			'Order Request Date
		strOrder = strOrder & UNIConvDate(Request("txtAccDate_Sub2")) & gColSep		'Order Promise Date
		strOrder = strOrder & "0" & gColSep											'Order Flag - Customer Order : 0
		strOrder = strOrder & "" & gColSep											'Request("txtCustomer")
	
		strOrder = strOrder & Trim(Request("txtItemCode")) & gColSep				'Part ID
		strOrder = strOrder & UNIConvNum(Request("txtAccQty_Sub2"),0) & gColSep		'Part Qty
		strOrder = strOrder & Trim(Request("txtTrackingNO")) & gRowSep				'Tracking No

	End Select

'== Global ����(DBLoginID, DBLoginPwd)�� incServer.asp�� ����Ǹ� �Ʒ� �ּ�ó���� ������ ����.. ��..
	
	'pis test
	'ProjModDate = objCon.ModifyOrder(strPlantCd,strOrder,gDBServer,gDatabase,gDBLoginID,gDBLoginPwd,strAPSHost,strAPSPort)

'==	ProjModDate = objCon.ModifyOrder(strPlantCd,strOrder,gDBServer,gDatabase,gDBLoginID,gDBLoginPwd,APSHOST,APSPORT)
			   
	If Err.Number <> 0 Then
		Set objCon = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)			'��:
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If

	If objCon.ErrorNo <> 0 Then
		Call ServerMesgBox(objCon.ErrorText, vbCritical, I_MKSCRIPT)              
		Set objCon = Nothing
		Response.End 
	End If


	Dim arrProjModVal
	arrProjModVal = Split(ProjModDate, gRowSep)

%>
<Script Language=vbscript>
	With parent.frm1
		<% '������(OLD)���������� ���� %>
		.txtBeforeChangeDate.value	= .txtAccDate_All.value
		<% '������(NEW)���������� ���� %>
		.txtAfterChangeDate.value	= "<%=UNIDateClientFormat(arrProjModVal(0))%>"
		parent.DbCTPSaveOk										
	End With
</Script>
<%

	'Set objCon = Nothing
	Response.End 	


Case CStr(UID_M0001)																'��: ���� ��û�� ���� 
									
    
    Err.Clear																		

    'Set pS31141 = Server.CreateObject("S31141.S31141SoSchdLine")
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set pS31141 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If

    '-----------------------
    'Data manipulate area
    '-----------------------
    'pis test
    'pS31141.ImpSSoHdrSoNo = TRIM(Request("txtSoNo"))
    'pS31141.ImpSSoDtlSoSeq = TRIM(Request("txtSoSeq"))
    'pS31141.ImpBItemItemCd = TRIM(Request("txtItemCode"))

	'-----------------------
	'Com Action Area
	'-----------------------
	'pS31141.CommandSent = "QUERY"
	'pS31141.ServerLocation = ggServerIP
    'pS31141.ComCfg = gConnectionString
	'pS31141.Execute

    If Err.Number <> 0 Then
		Set pS31141 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If

	'-----------------------
	'DB Error
	'-----------------------
    If Not (pS31141.OperationStatusMessage = MSG_OK_STR) Then
		Call DisplayMsgBox(pS31141.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		Call DisplayMsgBox(pS31141.ExpEabSqlCodeSqlcode, vbOKOnly, "", "", I_MKSCRIPT)
		Set pS31141 = Nothing
		Response.End 
    End If

	'-----------------------
	'Result data display area
	'----------------------- 
%>
<Script Language=vbscript>
	With parent.frm1

		.txtCtpCDFlag.value = "<%=pS31141.ExpGubunIefSuppliedSelectChar%>"
		.txtCtpSeq.value	= "<%=pS31141.ExpSSoDtlCtpTimes%>"

		Call parent.DisplayFlag(parent.Div1)

		If parent.UniCDbl(.txtCtpSeq.value) = 0 Then
			parent.DbCTPQuery
			'parent.CTPKaraCalc
		Else

			Select Case UCase(Trim(.txtCtpCDFlag.value))
			Case UCase("C")		'�����ϰ�� 

				'Call parent.DisplayFlag(parent.Div2)
				'.txtAccDate_All_Com.value	= "<%=UNIDateClientFormat(pS31141.ExpItemSSoSchdPromiseDt(1))%>"
				'.txtAccQty_All_Com.value	= "<%=pS31141.ExpItemSSoSchdCfmQty(1)%>"

				.txtAccDate_All.value	= "<%=UNIDateClientFormat(pS31141.ExpItemSSoSchdPromiseDt(1))%>"
				.txtAccQty_All.value	= "<%=UNINumClientFormat(pS31141.ExpItemSSoSchdCfmQty(1),ggQty.DecPoint,0)%>"
				.txtAccDate_Sub1.value	= ""
				.txtAccQty_Sub1.value	= ""
				.txtAccDate_Sub2.value	= ""
				.txtAccQty_Sub2.value	= ""

				.rdoSelect_All.checked = True
				Call parent.ggoOper.SetReqAttr(.rdoSelect_Sub, "Q")

			Case UCase("D")		'�����ϰ�� 

				'Call parent.DisplayFlag(parent.Div3)
				'.txtAccDate_Sub1_Div.value	= "<%=pS31141.ExpItemSSoSchdPromiseDt(1)%>"
				'.txtAccQty_Sub1_Div.value	= "<%=pS31141.ExpItemSSoSchdCfmQty(1)%>"
				'.txtAccDate_Sub2_Div.value	= "<%=pS31141.ExpItemSSoSchdPromiseDt(2)%>"
				'.txtAccQty_Sub2_Div.value	= "<%=pS31141.ExpItemSSoSchdCfmQty(2)%>"

				.txtAccDate_All.value	= ""
				.txtAccQty_All.value	= ""
				.txtAccDate_Sub1.value	= "<%=UNIDateClientFormat(pS31141.ExpItemSSoSchdPromiseDt(1))%>"
				.txtAccQty_Sub1.value	= "<%=UNINumClientFormat(pS31141.ExpItemSSoSchdCfmQty(1),ggQty.DecPoint,0)%>"
				.txtAccDate_Sub2.value	= "<%=UNIDateClientFormat(pS31141.ExpItemSSoSchdPromiseDt(2))%>"
				.txtAccQty_Sub2.value	= "<%=UNINumClientFormat(pS31141.ExpItemSSoSchdCfmQty(2),ggQty.DecPoint,0)%>"

				.rdoSelect_Sub.checked = True
				Call parent.ggoOper.SetReqAttr(.rdoSelect_All, "Q")

			End Select

		End If

	End With
</Script>
<%					

    'Set pS31141 = Nothing															'��: Unload Comproxy
	Response.End																	'��: Process End


Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
									
    Err.Clear																		

    'Set pS31141 = Server.CreateObject("S31141.S31141SoSchdLine")
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set pS31141 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If

    '-----------------------
    'Data manipulate area
    '-----------------------
    pS31141.ImpSSoHdrSoNo = Trim(Request("txtSoNo"))
    pS31141.ImpSSoDtlSoSeq = Trim(Request("txtSoSeq"))
    pS31141.ImpBItemItemCd = Trim(Request("txtItemCode"))
    pS31141.ImpSWksUserUserId = Trim(gUsrId)

	Select Case Trim(Request("txtRadioFlg"))
	Case "A"
		If Len(Trim(Request("txtAccDate_All"))) Then pS31141.ImpItemSSoSchdPromiseDt(1) = UNIConvDate(Trim(Request("txtAccDate_All")))
		pS31141.ImpItemSSoSchdCfmBaseQty(1) = UNIConvNum(Trim(Request("txtAccQty_All")),0)
		pS31141.ImpIefSuppliedCount = "1"
	Case "S"
		If Len(Trim(Request("txtAccDate_Sub1"))) Then pS31141.ImpItemSSoSchdPromiseDt(1) = UNIConvDate(Trim(Request("txtAccDate_Sub1")))
		pS31141.ImpItemSSoSchdCfmBaseQty(1) = UNIConvNum(Trim(Request("txtAccQty_Sub1")),0)
		If Len(Trim(Request("txtAccDate_Sub2"))) Then pS31141.ImpItemSSoSchdPromiseDt(2) = UNIConvDate(Trim(Request("txtAccDate_Sub2")))
		pS31141.ImpItemSSoSchdCfmBaseQty(2) = UNIConvNum(Trim(Request("txtAccQty_Sub2")),0)
		pS31141.ImpIefSuppliedCount = "2"								'ī������ ������ 
	End Select
    		                

	'-----------------------
	'Com Action Area
	'-----------------------
	pS31141.CommandSent = "SAVE"
	pS31141.ServerLocation = ggServerIP
    pS31141.ComCfg = gConnectionString
	pS31141.Execute

    If Err.Number <> 0 Then
		Set pS31141 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If

	'-----------------------
	'DB Error
	'-----------------------
    If Not (pS31141.OperationStatusMessage = MSG_OK_STR) Then
		Call DisplayMsgBox(pS31141.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		Set pS31141 = Nothing
		Response.End 
    End If

	'-----------------------
	'Result data display area
	'----------------------- 
%>
<Script Language=vbscript>
	With parent	
		.frm1.txtExitFlag.value = .CTPAccept
		.DbSaveOk
	End With
</Script>
<%					

    'Set pS31141 = Nothing																	'��: Unload Comproxy
	Response.End																			'��: Process End


Case CStr("CTPCancel")														'��: CTPQuery ��ȸ ��û�� ���� 

	Dim CancelVal

    '=======================
    'APS�� CTP Cancel Call
    '=======================
    Err.Clear                                                               

    'Set objCon = Server.CreateObject("uniAPS.APSConnect")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set objCon = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If

	strOrder = ""

	'-----------------------
    'Connection APS Server
    '-----------------------
	Select Case Request("txtRadioFlg")
	Case "A"
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & gRowSep				'Order ID - Distnct

	Case Else
		'--First Date AcceptOrder
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & "-1" & gColSep		'Order ID - Distnct

		'--Second Date AcceptOrder
		strOrder = strOrder & Trim(Request("txtSoNo")) _
					& "-" & Trim(Request("txtSoSeq")) & "-2" & gRowSep		'Order ID - Distnct

	End Select

'== Global ����(APSHOST,APSPORT : ����� ������ ���)�� incServer.asp�� ����Ǹ� �Ʒ� �ּ�ó���� ������ ����.. ��..
	
	'pis test	
	'CancelVal = objCon.CancelOrder(strPlantCd,strOrder,gDBServer,gDatabase,gDBLoginID,gDBLoginPwd,strAPSHost,strAPSPort)

'==	CancelVal = objCon.CancelOrder(strPlantCd,strOrder,gDBServer,gDatabase,gDBLoginID,gDBLoginPwd,APSHOST,APSPORT)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                                 '��:
	   Set objCon = Nothing																	'��: ComProxy UnLoad
	   Response.End																				'��: Process End
	End If
	
	If CancelVal = False Then
		Call ServerMesgBox("CTP Cancel Error", vbCritical, I_MKSCRIPT)              
		Set objCon = Nothing
		Response.End 
	End If
	
    '=======================
    'APS�� CTP Cancel�� �������ϰ�� ������ SoSchdLine Pad Call
    '=======================
    Err.Clear																		

    'Set pS31141 = Server.CreateObject("S31141.S31141SoSchdLine")
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set pS31141 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If

    '-----------------------
    'Data manipulate area
    '-----------------------
    pS31141.ImpSSoHdrSoNo = Trim(Request("txtSoNo"))
    pS31141.ImpSSoDtlSoSeq = Trim(Request("txtSoSeq"))
    pS31141.ImpBItemItemCd = Trim(Request("txtItemCode"))

	'-----------------------
	'Com Action Area
	'-----------------------
	pS31141.CommandSent = "CANCEL"
	pS31141.ServerLocation = ggServerIP
    pS31141.ComCfg = gConnectionString
	pS31141.Execute

    If Err.Number <> 0 Then
		Set pS31141 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If

	'-----------------------
	'DB Error
	'-----------------------
    If Not (pS31141.OperationStatusMessage = MSG_OK_STR) Then
		Call DisplayMsgBox(pS31141.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
		Call DisplayMsgBox(pS31141.ExpEabSqlCodeSqlcode, vbOKOnly, "", "", I_MKSCRIPT)
		Set pS31141 = Nothing
		Response.End 
    End If

	'-----------------------
	'Result data display area
	'----------------------- 
%>
<Script Language=vbscript>
	With parent
		MsgBox "CTP Cancel�� �Ϸ�Ǿ����ϴ�.", vbInformation, "<%=gLogoName%>"
		.frm1.txtExitFlag.value = .CTPCancel

		If .UNICDbl(.frm1.txtCtpSeq.value) > 0 Then
			.Self.Returnvalue	= .CTPModify										<%'��: �����Ͻ� ó�� ASP �� ���� %>
		Else
			.Self.Returnvalue	= .CTPAccept										<%'��: �����Ͻ� ó�� ASP �� ���� %>
		End If

		.CancelClickOK
	End With
</Script>
<%					

    'Set pS31141 = Nothing															'��: Unload Comproxy
	Response.End																	'��: Process End

End Select

'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
%>
<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>
</SCRIPT>
