
delete from STATEMENTS
where id ='p4112mb26_lko391'

insert STATEMENTS
 	(id,
	stype,
	module,
	def_type,
	head,
	body,
	tail,
	description)
 values
	('p4112mb26_lko391',
	'2',
	'PP',
	'S',
    ' SELECT  top 51 a.opr_no, e.job_cd, f.minor_nm, c.wc_cd, c.wc_nm,'+
    '               b.item_cd, b.item_nm, b.spec,'+
    '               a.req_qty, a.base_unit, a.issued_qty,'+
    '               a.req_dt, a.tracking_no,'+
    '               a.sl_cd,  d.sl_nm,'+
    '               a.resv_status, dbo.ufn_GetCodeName(''P1017'', a.resv_status) as Resv_Desc,'+
    '               a.issue_mthd, dbo.ufn_GetCodeName(''P1016'', a.issue_mthd) as Issue_Mthd_Desc,'+
    '               a.req_no, a.seq, a.prodt_order_no, e.order_status, e.inside_flg'+
    ' from p_reservation a,'+
    '     b_item b,'+
    '     p_work_center c,'+
    '     b_storage_location d,'+
    '     p_production_order_detail e,'+
    '     b_minor f',
    ' a.child_item_cd = b.item_cd and a.wc_cd = c.wc_cd'+
    ' and a.sl_cd =d.sl_cd and a.prodt_order_no =e.prodt_order_no and a.opr_no = e.opr_no'+
    ' and (e.job_cd *= f.minor_cd and f.major_cd = ''P1006'')'+
    ' , c.plant_cd = ? , a.prodt_order_no = ?',
	'order by a.opr_no',
	'List Production Order Detail')