SELECT  /*+ PARALLEL(16) */ re.*,
            GREATEST(ne.dep, ne.sdo, ne.cr, ne.pkg, ne.ved, ne.do, ne.txn) AS ACT_NEW_V3,
            ne.dep AS deposit,
            ne.sdo AS SDO_tek,
            ne.cr AS KP,
            ne.pkg AS pocket,
            ne.ved AS VED,
            ne.do AS DO,
            ne.txn AS Deb_trans,
            ne.sum_dep,
            ne.sum_sdo,
            ne.sum_od,
            ne.sum_pkg,
            ne.count_ved,
            ne.sum_fx,
            ne.count_fx,
            ne.sum_do,
            ne.count_txn,
            CASE WHEN ROW_NUMBER() OVER (PARTITION BY re.tax_id, re.branch_code ORDER BY ne.bin) = 1 THEN 1 ELSE 2 END AS Flag
FROM 
(select  /*+ PARALLEL(16) */
            case when re.branch_code = '***/01' then '**' ELSE re.branch_code end branch_code,
            case when re.filial_nm='Филиал в г.*****' then 'Филиал в г.******' 
                 when re.branch_code = '***/01' then 'Филиал в г.*****'
                 else re.filial_nm end filial_nm,
            re.tax_id tax_id,
            re.branch_code branch_code,
            re.client_name client_name,
            re.client_id client_id,
            re.OKED OKED,
            re.segment_1 segment_1,
            re.segment_2 segment_2,
            re.open_date open_date,
            re.REPORT_DATE REPORT_DATE,
            re.last_pos last_pos,
            re.cust_type cust_type,
            re.bal_kzt bal_kzt,
            re.country_ country,
            re.operator_id operator_id,
            re.Est_schet Have_acct,
            coalesce (d.k1,0) K1,
            coalesce (d.k2,0) K2,
            coalesce (o.count_block,0) Blok,
            case when coalesce(c1.date_in,to_date('01012000','dd.mm.yyyy'))>= trunc(re.REPORT_DATE)-90 then 1 else 0 end NTB,
            c1.date_in
from (
        select t1.branch_code,
               fl.filial_nm,
               t.tax_id,
               t1.branch_code,
               t.client_name,
               cc.client_id client_id1,
               t.client_id,
               s1.xattr_value OKED,
               case when coalesce (h.segm,t.code_value) ='*' then '*******'
                    when coalesce (h.segm,t.code_value) in('*1','*2') then '***-***'
                    when coalesce (h.segm,t.code_value) in '*3','*4','*5' then '**-**-**'
                    when coalesce (h.segm,t.code_value) in '*6','*7','*8' then '*****-*'
                when coalesce (h.segm,t.code_value) in '*8','*9'then '***-**'
                  when coalesce (h.segm,t.code_value) = '*10' then '**-*'
               else coalesce (h.segm,t.code_value) end as segment_1,
               coalesce (h.segm,t.code_value) as segment_2,
               max(t1.open_date) open_date,
               t.report_date,
               max(t1.last_pos) last_pos,
               case when cc.cust_stat in ('***','**') then '**' else '**' end as cust_type,
               -sum(t.bal_kzt) as bal_kzt,
               cc.region_code country_,
               max(t1.operator_id) operator_id,
               '1' as Est_schet
        from source_t1 t
        left join source_t2 s on s.file_name='***-***' and s.code_value=t.client_id and s.code='***'
        left join source_t3 cc on to_char(cc.client_id)=s.surrogate
        left join source_t2 s1 on s1.surrogate=to_char(cc.client_id) and s1.code='***1' and s1.file_name='***-***1'
        left join(select distinct
                         t.report_date,
                         t.client_id,
                         t.branch_code,
                         t.tax_id,
                         first_value(t.open_date) over (partition by t.client_id,t.branch_code,t.report_date order by t.open_date) as open_date,
                         first_value(t.branch_code) over (partition by t.client_id,t.branch_code,t.report_date order by t.open_date) as branch_code,
                         first_value(t.acct_operator_id) over (partition by t.client_id,t.branch_code,t.report_date order by t.open_date) as operator_id,
                         nvl(case when first_value(t.last_pos) over (partition by t.client_id,t.branch_code,t.report_date order by nvl(t.last_pos,'01.01.1900') desc)='01.01.1900' then null else first_value(t.last_pos) over (partition by t.client_id,t.branch_code,t.report_date order by nvl(t.last_pos,'01.01.1900') desc) end
                         ,case when t.open_date <= t.report_date then first_value(t.open_date) over (partition by t.client_id,t.branch_code,t.report_date order by t.open_date desc) end) as last_pos
                  from source_t1 t
                  where t.report_date = (select max(report_date) from source_t1) and t.end_date is null
                  and rtrim(ltrim(t.acct_operator_id)) not in ('*2*','******2','****1*','D****','****S','S****')
             ) t1 on t1.client_id=t.client_id and t1.tax_id=t.tax_id and t1.branch_code=t.branch_code and t1.report_date=t.report_date
        left join source_t4 fl on fl.branch_code=t.branch_code
        LEFT JOIN fd.t_dict_segment_history h
             ON h.is_actual = 4876871
             AND h.cust_cat = '*'
             AND h.tax_id = t.tax_id
             AND h.unk = t.client_id
        where t.report_date = (select max(report_date) from source_t1) and rtrim(ltrim(t.acct_operator_id)) not in ('*2*','******2','****1*','D****','****S','S****')
        and t.end_date is null
        group by t.client_id,s1.xattr_value,t.tax_id,t.client_name,t.report_date,cc.client_id,t1.branch_code,t1.branch_code,coalesce (h.segm,t.code_value),fl.filial_nm,cc.region_code,case when cc.cust_stat in ('**','**2') then '**' else 'ЮЛ' end
              ) re
left join
(
select nvl(d1.client_id, d2.client_id) client_id, nvl(d1.branch_code,d2.branch_code) branch_code, nvl(d1.balance_amt,0) k2, nvl(d2.balance_amt,0) k1
from (
select c.client_id, a.branch_code, max(pos.balance_amt) balance_amt
FROM source_t5  pos
join source_t6 a
     on a.cust_cat = '*'
     and pos.since = (select max(since) from source_t5)
     and pos.acct = a.acct
     and pos.currency_code = a.currency_code
     and a.misc##2 = '********'
     and pos.balance_amt > 0
join source_t3 c
     on a.client_id = c.client_id
group by c.client_id, a.branch_code ) d1
full join (
select c.client_id,a.branch_code, sum(pos.balance_amt) balance_amt
FROM source_t5  pos
join source_t6 a
     on a.cust_cat = '*'
     and pos.since = (select max(since) from source_t5)
     and pos.acct = a.acct
     and pos.currency_code = a.currency_code
     and a.misc##2 = '*****-****'
     and pos.balance_amt > 0
join source_t3 c
     on a.client_id = c.client_id
group by c.client_id, a.branch_code ) d2 on d2.client_id = d1.client_id and d2.branch_code = d1.branch_code
) d on d.client_id = re.client_id1 and d.branch_code = re.branch_code
left join
(
select c.client_id,a.branch_code, count(distinct(b.txt##4))count_block
from source_t3 c
join source_t6 a
     on a.cust_cat = '*'
     and a.client_id = c.client_id
     and c.class_code in ('***-***I','***-***C')
     and (a.close_date >= trunc(sysdate) or a.close_date is null)
join source_t7 B
     on B.SURROGATE = A.ACCT ||','|| (CASE WHEN A.currency_code = ' ' THEN '' ELSE A.currency_code END)
     AND FILE_NAME = 'ACCT'
     AND (B.End_Time is null or b.end_time > trunc(sysdate))
     and b.block_type != '*****1'
group by c.client_id,a.branch_code
) o on o.client_id = re.client_id1 and o.branch_code = re.branch_code
left join 
     (
     select c.tax_id, c.date_in
     from source_t3 c
     where c.class_code in ('***-***C','***-***I')
     and c.tax_id != '*******'
     ) c1 on c1.tax_id = re.tax_id) re
LEFT JOIN 
     source_t8 ne
     ON ne.bin = re.tax_id
     AND ne.bin != '*******'
     AND ne.report_date = re.REPORT_DATE