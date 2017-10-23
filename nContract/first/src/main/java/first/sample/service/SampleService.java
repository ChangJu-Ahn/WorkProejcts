package first.sample.service;

import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;


public interface SampleService {

	List<Map<String, Object>> selectBoardList(Map<String, Object> map) throws Exception;

	void insertBoard(Map<String, Object> map, HttpServletRequest request) throws Exception;

	List<Map<String, Object>> openPopupList(Map<String, Object> map) throws Exception;
	
	void modifyBoard(Map<String, Object> map, HttpServletRequest request) throws Exception;

	List<Map<String, Object>> openContractPopupList(Map<String, Object> commandMap) throws Exception;

	void endBoard(Map<String, Object> map, HttpServletRequest request) throws Exception;

	Map<String, Object> selectBoardDetail(Map<String, Object> map) throws Exception;

	void updateBoard(Map<String, Object> map, HttpServletRequest request) throws Exception;

	Map<String, Object> selectBoardUpdateDetail(Map<String, Object> map);

	Map<String, Object> openHstList(Map<String, Object> commandMap);

	List<Map<String, Object>> openGubunPopupList(Map<String, Object> commandMap);

	void updateCotents(Map<String, Object> map);

	void sendEmail(Map<String, Object> map);

	List<Map<String, Object>> selectBoardSearchList(Map<String, Object> map) throws Exception;

	List<Map<String, Object>> selectAdminCode(Map<String, Object> map) throws Exception;

	void updateGrid(Map<String, Object> map) throws Exception;

	void addGrid(Map<String, Object> map) throws Exception;

	void delGrid(Map<String, Object> map) throws Exception;

	Map<String, Object> NtotalContract(Map<String, Object> map) throws Exception;

	List<Map<String, Object>> NperiodContract(Map<String, Object> map) throws Exception;

	List<Map<String, Object>> AperiodContract(Map<String, Object> map) throws Exception;

	List<Map<String, Object>> AtotalContract_Simple() throws Exception;
	
	List<Map<String, Object>> AtotalContract_Detail() throws Exception;

	List<Map<String, Object>> selectBoxList(Map<String, Object> map) throws Exception;

	List<Map<String, Object>> getAllData(Map<String, Object> map) throws Exception;

	List<Map<String, Object>> selectUserList(Map<String, Object> map) throws Exception;
	
	List<Map<String, Object>> getStandardCode() throws Exception;

	void updateUserGrid(Map<String, Object> map) throws Exception;
	
	void addUserGrid(Map<String, Object> map) throws Exception;
	
	void delUserGrid(Map<String, Object> map) throws Exception;
	
	void deleteContract(Map<String, Object> map) throws Exception;

	void deleteBoardHst(Map<String, Object> map)  throws Exception;
	
	void updateBoardHst(Map<String, Object> map, HttpServletRequest request)  throws Exception;

	Map<String, Object> selectBoardUpdateHstDetail(Map<String, Object> map) throws Exception;

	void initUserInfo(Map<String, Object> map) throws Exception;
	
	boolean updateUserInfo(Map<String, Object> map) throws Exception;	
	
	void setLoginHistoryWrite(Map<String, Object> map) throws Exception;
}
