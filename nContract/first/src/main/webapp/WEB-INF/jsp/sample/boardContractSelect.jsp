<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
//뒤로가기 버튼을 사용하기 위해 캐쉬 제거
    response.setDateHeader("Expires", 0);
    response.setHeader("Pragma", "no-cache");
    if(request.getProtocol().equals("HTTP/1.1")) {
        response.setHeader("Cache-Control", "no-cache");
    }
%>
<html>
<head>
	<%@ include file="/WEB-INF/include/include-header.jspf" %>
</head>
<body>
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
	<H2>체결 계약서 이력</H2>
    <table class="board_view">
	    <colgroup>
			<col width="20%"/>
			<col width="80%"/>
		</colgroup>
		<tbody>
			<tr>
				<th scope="row">계약 번호</th>
				<td colspan="5">${map.CONTRACT_NO }</td>
			</tr>
	   		<tr>
				<th scope="row">계약 이력</th>
					<td>
						<c:forEach var="row" items="${h_list }" varStatus="var">
						 <div>
	                       <input type="hidden" id="HST_SEQ" value="${row.HST_SEQ }">
	                       <a href="#this" name="CONTRACT_NM">${row.HST_SEQ }. ${row.CONTRACT_NM }</a>
	                       <c:choose>
	                       		<c:when test="${row.EXPIRE_FLAG == 'Y' && row.MODIFY_FLAG == 'Y'}">
	                       			(해지)
	                       		</c:when>
	                       		<c:when test="${row.EXPIRE_FLAG == 'N' && row.MODIFY_FLAG == 'Y'}">
	                       			(변경)
	                       		</c:when>
	                       		<c:otherwise>
	                       			(등록)
	                       		</c:otherwise>
	                       </c:choose>
	                     </div>
	                   	</c:forEach>
	               	</td>
				</tr>
			</tbody>
    </table>
   	<table class="board_list">
       <colgroup>
        <col width="10%"/>
		<col width="5%"/>
		<col width="15%"/>
		<col width="5%"/>
		<col width="20%"/>
		<col width="15%"/>
		<col width="15%"/>
		<col width="5%"/>
		<col width="5%"/>
	</colgroup>
	<thead>
		<tr>
			<th scope="col">계약 번호</th>
			<th scope="col">사업부</th>
			<th scope="col">고객사(업체)1 / 고객사(업체)2</th>
			<th scope="col">구분</th>
			<th scope="col">계약서 명</th>
			<th scope="col">목적 사업(제품)</th>
			<th scope="col">효력발생일 / 종료일</th>
			<th scope="col">자동연장</th>
			<th scope="col">만료</th>
		</tr>
	</thead>
	<tbody>
    </tbody>
    </table>
    
    <%@ include file="/WEB-INF/include/include-body.jspf" %>
    <script type="text/javascript">
    $(document).ready(function(){
    	
    	
    });
    </script>
</body>
</html>