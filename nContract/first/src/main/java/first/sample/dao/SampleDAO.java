package first.sample.dao;

import java.util.List;
import java.util.Map;


import org.springframework.stereotype.Repository;

import first.common.dao.AbstractDAO;

@Repository("sampleDAO")
public class SampleDAO extends AbstractDAO{

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectBoardList(Map<String, Object> map) throws Exception{
		return (List<Map<String, Object>>)selectPagingList("sample.selectBoardList", map);
		/*return (List<Map<String, Object>>)selectList("sample.selectBoardList", map);*/
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectBoardBusorList(Map<String, Object> map) throws Exception{
		return (List<Map<String, Object>>)selectPagingList("sample.selectBoardBusorList", map);
		/*return (List<Map<String, Object>>)selectList("sample.selectBoardList", map);*/
	}
	
	public void insertBoard(Map<String, Object> map) {
		insert("sample.insertBoard", map);
	}
	
	public void insertHeaderBoard(Map<String, Object> map) {
		insert("sample.insertHeaderBoard", map);
	}
	
	public void insertDetailBoard(Map<String, Object> map) {
		insert("sample.insertDetailBoard", map);
	}
	
	public void insertHistoryBoard(Map<String, Object> map) {
		insert("sample.insertHistoryBoard", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> openPopupList(Map<String, Object> map) throws Exception{
		return (List<Map<String, Object>>)selectList("sample.openPopupList1", map);
	}

	public void insertFile(Map<String, Object> map) throws Exception{
		insert("sample.insertFile", map);
	}

	public void insertHstFile(Map<String, Object> map) throws Exception{
		insert("sample.insertHstFile", map);
	}
	
	
	public void modifyBoard(Map<String, Object> map) throws Exception{
		update("sample.modifyBoard", map);		
	}

	public void modifyFile(Map<String, Object> map) throws Exception{
		update("sample.modifyFile", map);	
	}

	public void updateFileDel(Map<String, Object> map) throws Exception{
		update("sample.updateFileDel", map);
	}
	
	public void updateFileSeq(Map<String, Object> map) throws Exception{
		update("sample.updateFileSeq", map);
	}
	
	public void updateFileSeqToHeader(Map<String, Object> map) throws Exception{
		update("sample.updateFileSeqToHeader", map);
	}
	
	public void endBoard(Map<String, Object> map) throws Exception{
		update("sample.endBoard", map);
	}
	
	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> openContractPopupList(Map<String, Object> map) throws Exception{
		return (List<Map<String, Object>>)selectList("sample.openPopupList3", map);
	}
	
	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> openGubunPopupList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.openPopupList", map);
	}
	
	@SuppressWarnings("unchecked")
	public Map<String, Object> selectBoardDetail(Map<String, Object> map) throws Exception{
		return (Map<String, Object>) selectOne("sample.selectBoardDetail", map);
	}
	
	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> getStandardCode() throws Exception{
		return (List<Map<String, Object>>) selectList("sample.selectStandardCode");
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectFileList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.selectFileList", map);
	}

	public void updateBoard(Map<String, Object> map) {
		update("sample.updateBoard", map);
	}

	@SuppressWarnings("unchecked")
	public Map<String, Object> selectUpdateBoardDetail(Map<String, Object> map) {
		return (Map<String, Object>) selectOne("sample.selectBoardUpdateDetail", map);
	}
	
	@SuppressWarnings("unchecked")
	public Map<String, Object> selectUpdateBoardHstDetail(Map<String, Object> map) {
		return (Map<String, Object>) selectOne("sample.selectBoardUpdateHstDetail", map);
	}

	public void deleteFileList(Map<String, Object> map) {
		update("sample.deleteFileList", map);
	}

	public void updateFile(Map<String, Object> map) {
		update("sample.updateFile", map);
	}

	public void insertHst(Map<String, Object> map) {
		insert("sample.insertHst", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectHstList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.selectHstList", map);
	}

	@SuppressWarnings("unchecked")
	public Map<String, Object> selectBoardHstDetail(Map<String, Object> map) {
		return (Map<String, Object>) selectOne("sample.selectBoardHstDetail", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectHstFileList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.selectHstFileList", map);
	}

	public void updateContents(Map<String, Object> map) {
		update("sample.updateContents", map);
	}

	public void sendEmail(Map<String, Object> map) {
		update("sample.sendEmail", map);
		
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectBoardSearchList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectPagingList("sample.selectBoardSearchList", map);
	}
	
	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectBoardBusorSearchList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectPagingList("sample.selectBoardBusorSearchList", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectBoardAdminCode(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.selectBoardAdminCode", map);
	}

	public void updateGrid(Map<String, Object> map) {
		update("sample.updateGrid", map);
	}

	public void addGrid(Map<String, Object> map) {
		insert("sample.addGrid", map);
	}

	public void delGrid(Map<String, Object> map) {
		delete("sample.delGrid", map);
	}

	public void updateUserGrid(Map<String, Object> map) {
		update("sample.updateUserGrid", map);
	}

	public void addUserGrid(Map<String, Object> map) {
		insert("sample.addUserGrid", map);
	}

	public void delUserGrid(Map<String, Object> map) {
		delete("sample.delUserGrid", map);
	}
	
	public void deleteContract(Map<String, Object> map) {
		delete("sample.deleteContract", map);
	}
	
	public void updateContractHst(Map<String, Object> map) {
		update("sample.updateContractHst", map);
	}
	
	public void deleteContractHst(Map<String, Object> map) {
		delete("sample.deleteContractHst", map);
	}
	
	public void updateUserInfo_initial(Map<String, Object> map) {
		update("sample.updateUserInfo_initial", map);
	}
	
	public boolean updateUserInfo_update(Map<String, Object> map) {
		boolean TempResult;
		Integer cnt = (Integer)update("sample.updateUserInfo_update", map);
		
		//업데이트가 되었으면 카운트가 올라 감, 업데이트가 되었다는 건 비밀번호가 변경되었다는 의미로  return
		if(cnt > 0)
			TempResult = true;
		else  
			TempResult = false;
		
		return TempResult;
	}
	
	public void insertLoginHistory(Map<String, Object> map) {
		insert("sample.insertLoginHistory", map);
	}
	
	public void updateLoginHistory(Map<String, Object> map) {
		update("sample.updateLoginHistory", map);
	}
	
	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> NtotalContractGraph(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.NtotalContractGraph", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> NtotalContractGrid(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.NtotalContractGrid", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> NperiodContract(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.NperiodContract", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> AperiodContract(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.AperiodContract", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> AtotalContract_Detail() {
		return (List<Map<String, Object>>)selectList("sample.AtotalContract_Detail");
	}
	
	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> AtotalContract_Simple() {
		return (List<Map<String, Object>>)selectList("sample.AtotalContract_Simple");
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectBoxList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.selectBoxList");
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> getAllData(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.getAllData", map);
	}
	
	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> selectUserList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.selectUserList", map);
	}

	@SuppressWarnings("unchecked")
	public List<Map<String, Object>> getAllAdminCodeList(Map<String, Object> map) {
		return (List<Map<String, Object>>)selectList("sample.selectAdminCodeTotalList", map);
	}
}
