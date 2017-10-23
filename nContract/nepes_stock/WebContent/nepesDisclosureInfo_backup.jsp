<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<%@page import="java.beans.XMLEncoder"%>
<%@ page import="org.w3c.dom.*"%>
<%@ page import="org.xml.sax.*"%>
<%@ page import="java.util.*"%>
<%@ page import="java.io.*"%>
<%@ page import="java.net.*"%>
<%@ page import="javax.xml.parsers.*"%>
<%@ page import="javax.servlet.http.HttpServletResponse.*"%>
<%
	String geturl = "http://asp1.krx.co.kr/servlet/krx.asp.DisList4MainServlet?code=033640&gubun=K";
	String gettime = "";

	String xmlstr = "";
	int disInfo_lenth = 0;

	String line = "";
	String disInfo[][] = new String[10][4];

	String xml = "";

	try {
		URL url = new URL(geturl);
		URLConnection conn = url.openConnection();
		HttpURLConnection httpConnection = (HttpURLConnection) conn;
		InputStream is = null;
		InputStreamReader isr = null;

		is = new URL(geturl).openStream();
		isr = new InputStreamReader(is, "euc-kr");

		BufferedReader rd = new BufferedReader(isr, 400);

		StringBuffer strbuf = new StringBuffer();

		while ((line = rd.readLine()) != null) {

			strbuf.append(line);
		}

		//System.out.println("공시정보");
		//System.out.println(strbuf.toString().trim());

		DocumentBuilderFactory docFact = DocumentBuilderFactory.newInstance();
		docFact.setNamespaceAware(true);
		DocumentBuilder docBuild = docFact.newDocumentBuilder();

		Document doc = docBuild.parse(new InputSource(new StringReader(strbuf.toString())));
		doc.getDocumentElement().normalize();

		Element root = doc.getDocumentElement();

		NodeList disclosureMain = doc.getElementsByTagName("disclosureMain");

		NamedNodeMap disclosureMaininfo = disclosureMain.item(0).getAttributes();
		gettime = disclosureMaininfo.getNamedItem("querytime").getNodeValue();

		NodeList disinfo = doc.getElementsByTagName("disInfo");

		disInfo_lenth = disinfo.getLength();
		for (int i = 0; i < disInfo_lenth; i++) {

			NamedNodeMap Disinfo = disinfo.item(i).getAttributes();

			disInfo[i][0] = Disinfo.getNamedItem("distime").getNodeValue();
			disInfo[i][1] = Disinfo.getNamedItem("disTitle").getNodeValue();
			disInfo[i][2] = Disinfo.getNamedItem("disAcpt_no").getNodeValue();
			disInfo[i][3] = Disinfo.getNamedItem("submitOblgNm").getNodeValue();

		}

	} catch (Exception e) {

	}
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=EUC-KR">
<!-- jQuery -->
<script src="//code.jquery.com/jquery-1.11.3.min.js"></script>
<script src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
<script src="<c:url value='/js/common.js'/>" charset="utf-8"></script>
<script src="<c:url value='/js/menu.js'/>" charset="utf-8"></script>
<script src="<c:url value='/js/lobibox.min.js'/>"></script>
<link rel="StyleSheet" type="text/css" href="./css/default.css" />
<title>nepes 공시정보</title>
<script type="text/javascript">
	
</script>
</head>
<body style="margin: 0px;">
	<FORM name="krx" method="post">
		<div style="width: 710px;">
			<div style="width: 700px; float: left; padding: 0 5 0 0;">
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<caption class="dpn">현재시간</caption>
					<tr>
						<td height="26" align="center">
							<table width="96%" border="0" cellpadding="0" cellspacing="0">
								<caption class="dpn">현재시간 상세내용</caption>
								<tr>
									<th><span
										style="font-family: 굴림, 굴림체, 돋움, 돋움체; font-size: 9pt; color: #000000;"><strong>현재시간</strong>
											: <%=gettime%></span></th>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</div>
			<div style="width: 700px; float: left; padding: 0 5 5 0;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<caption class="dpn">경영공시</caption>
					<tr>
						<td height="2" class="line_color"></td>
					</tr>
					<tr>
						<td height="2" bgcolor="#FFFFFF"></td>
					</tr>
					<tr>
						<td>
							<table id="dcg_mngdiscls_01" cellpadding="0" cellspacing="1"
								bgcolor="#EBEBEB" border="0" width="100%">
								<caption class="dpn">경영공시 상세내용</caption>
								<thead>
									<tr>
										<th width="6%" height="30" align="center" class="hd_bgcolor"><span
											class="hd_cfs"><strong>번호</strong></span></th>
										<th width="18%" align="center" class="hd_bgcolor"><span
											class="hd_cfs"><strong>시간</strong></span></th>
										<th width="58%" align="center" class="hd_bgcolor"><span
											class="hd_cfs"><strong>공시제목</strong></span></th>
										<th width="18%" align="center" class="hd_bgcolor"><span
											class="hd_cfs"><strong>제출의무자</strong></span></th>
									</tr>
								</thead>
								<%
									if (disInfo_lenth > 0) {
								%>
								<%
									for (int i = 0; i < disInfo_lenth; i++) {
								%>
								<%
									if ((i % 2) == 0) {
								%>
								<tr>
									<td height="30" align="center" class="bd_bgcolor"><span
										class="bd_cfs"><%=disInfo_lenth - i%></span></td>
									<td align="center" class="bd_bgcolor"><span class="bd_cfs"><%=disInfo[i][0].substring(0, 4)%>/<%=disInfo[i][0].substring(4, 6)%>/<%=disInfo[i][0].substring(6, 8)%></span></td>
									<td class="bd_bgcolor"><a href="#"
										onclick="window.open('http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno=<%=disInfo[i][2]%>','공시상세보기','width=1200,height=800,top=100,left=350');"><span
											style="padding: 0 0 0 5" class="bd_cfs"><%=disInfo[i][1]%></span></a></td>
									<td class="bd_bgcolor"><span style="padding: 0 0 0 5"
										class="bd_cfs"><%=disInfo[i][3]%></span></td>
								</tr>
								<%
									} else {
								%>
								<tr>
									<td height="30" align="center" bgcolor="#F5F5F5"><span
										class="bd_cfs"><%=disInfo_lenth - i%></span></td>
									<td align="center" bgcolor="#F5F5F5"><span class="bd_cfs"><%=disInfo[i][0].substring(0, 4)%>/<%=disInfo[i][0].substring(4, 6)%>/<%=disInfo[i][0].substring(6, 8)%></span></td>
									<td bgcolor="#F5F5F5"><a href="#"
										onclick="window.open('http://kind.krx.co.kr/common/disclsviewer.do?method=search&acptno=<%=disInfo[i][2]%>','공시상세보기','width=1200,height=800,top=100,left=350');"><span
											style="padding: 0 0 0 5" class="bd_cfs"><%=disInfo[i][1]%></span></a></td>
									<td bgcolor="#F5F5F5"><span style="padding: 0 0 0 5"
										class="bd_cfs"><%=disInfo[i][3]%></span></td>
								</tr>
								<%
									}
								%>
								<%
									}
								%>
								<%
									} else {
								%>
								<tr>
									<td colspan="4">데이터가 없습니다.</td>
								</tr>
								<%
									}
								%>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table cellspacing="0" cellpadding="0" border="0">
								<tr>
									<td height="15"></td>
								</tr>
								<tr>
									<td><img src="./img/wz_ment_1.gif"
										alt="본 정보는 오류가 발생하거나 지연 될 수 있습니다. 제공된 정보에 의한 투자결과에 대한 법적인 책임을 지지 않습니다." /></td>
								</tr>

							</table>
						</td>
					</tr>
					<tr>
						<td align="right" height="30"><a
								href="http://dart.fss.or.kr/html/search/SearchCompany_M2.html?textCrpNM=033640"
								style="color: Blue;" target="_blank" >더 많은
									정보 보기 ></a></td>
					</tr>
				</table>
			</div>
		</div>
	</FORM>
</body>
</html>