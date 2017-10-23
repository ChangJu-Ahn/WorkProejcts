-- 수정(20080109::hanc) : b.BASE_ITEM_CD = x.ITEM_CD --> b.BASE_ITEM_CD *= x.ITEM_CD
-- 수정(20080118::hanc) : i.opr_no, i.wc_cd, j.wc_nm --> '''' opr_no, '''' wc_cd, '''' wc_nm
-- 수정(20080226::hanc) : '''' opr_no, '''' wc_cd, '''' wc_nm --> i.opr_no, i.wc_cd, j.wc_nm 
-- 수정(20080227::hanc) : i.opr_no, i.wc_cd, j.wc_nm --> '''' opr_no, '''' wc_cd, '''' wc_nm  (이유 : 이렇게 해주어야 2줄 중복을 막을 수 있음)

-- 제조오더등록(소요량조정포함)(S)
delete from statements where id='p4112mb21_ko441'

insert into statements(
id,stype,module,def_type,head,body,tail,description) values (
'p4112mb21_ko441',2,'PP','S',
'  select distinct top 101 a.*, b.item_nm, b.spec, c.sl_nm, d.valid_from_dt, d.valid_to_dt, d.order_unit_mfg, d.order_lt_mfg, d.fixed_mrp_qty,         d.min_mrp_qty, d.max_mrp_qty, d.round_qty, d.scrap_rate_mfg,         dbo.ufn_GetCodeName(''P1012'', d.mps_mgr) mps_mgr,         dbo.ufn_GetCodeName(''P1011'', d.mrp_mgr) mrp_mgr,         dbo.ufn_GetCodeName(''P1015'', d.prod_mgr) prod_mgr,         f.mrp_run_no, e.parent_order_no , e.parent_opr_no, b.item_group_cd, g.item_group_nm,         a.cost_cd, h.cost_nm,         '''' opr_no, '''' wc_cd, '''' wc_nm, b.BASE_ITEM_CD, x.item_nm BASE_ITEM_NM  from p_production_order_header a, b_item b, b_storage_location c, b_item_by_plant d,      p_rework_order_history e, p_planned_order f, b_item_group g, b_cost_center h,      p_production_order_detail i, p_work_center j, b_item x	',
'  a.prodt_order_no = i.prodt_order_no and i.wc_cd = j.wc_cd and a.plant_cd = j.plant_cd and a.item_cd = b.item_cd   and a.plant_cd = d.plant_cd   and a.item_cd = d.item_cd   and a.plant_cd = f.plant_cd and a.plan_order_no = f.plan_order_no and a.sl_cd = c.sl_cd and a.prodt_order_no *= e.rework_order_no and b.item_group_cd *= g.item_group_cd and a.plant_cd *= h.plant_cd and a.cost_cd *= h.cost_cd AND   b.BASE_ITEM_CD *= x.ITEM_CD , a.plant_cd = ?, a.plan_start_dt >= ?, a.plan_start_dt <= ?, a.item_cd = ?, a.tracking_no = ?, a.prodt_order_no >= ?, a.prodt_order_type = ?, a.order_status = ?, f.mrp_run_no = ?, ?, ?	',
'  order by a.prodt_order_no ',
'List Production Order Header '
)
