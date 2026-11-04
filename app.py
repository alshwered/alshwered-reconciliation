import io, re, os
from datetime import date
import pandas as pd
import streamlit as st

st.set_page_config(page_title='مطابقة الشويرد', page_icon='📊', layout='wide')
st.markdown('''<style>
html, body, [class*="css"] {direction: rtl; text-align: right;}
[data-testid="stDataFrame"] {direction: rtl;}
.block-container {padding-top: 1.5rem;}
</style>''', unsafe_allow_html=True)

PROVIDERS = ['مدفوع','مس باي','تابي','تمارا','إمكان','زد باي','أمازون']
ZID_COLS = {
    'order':'id','name':'customer_name','phone':'customer_mobile','method':'payment_method',
    'payment_status':'payment_status','amount':'total','reference':'transaction_reference',
    'order_date':'added_at (Asia/Riyadh)','updated':'last_update_at (Asia/Riyadh)',
    'split1_name':'split_payment_method_1_name','split1_amount':'split_payment_method_1_total',
    'split2_name':'split_payment_method_2_name','split2_amount':'split_payment_method_2_total',
}

def norm_phone(v):
    if pd.isna(v): return ''
    s = re.sub(r'\D','',str(v))
    if s.startswith('00966'): s=s[2:]
    if s.startswith('966'): s=s[3:]
    if s.startswith('05'): s=s[1:]
    if s.startswith('5') and len(s)==9: return '966'+s
    if len(s)==10 and s.startswith('05'): return '966'+s[1:]
    return s

def num(v):
    return pd.to_numeric(v, errors='coerce').fillna(0)

def read_zid(file):
    raw = pd.read_excel(file, dtype={'customer_mobile':str})
    if 'id' not in raw.columns: raise ValueError('الملف لا يحتوي على رقم الطلب (id).')
    raw = raw[raw['id'].notna()].copy()
    raw['id'] = raw['id'].astype(str).str.replace(r'\.0$','',regex=True)
    # Zid exports continuation rows for additional products; retain one order-level row.
    orders = raw.drop_duplicates('id', keep='first').copy()
    for c in ['total','vat','shipping','discount','sub_totals','split_payment_method_1_total','split_payment_method_2_total']:
        if c in orders: orders[c] = pd.to_numeric(orders[c],errors='coerce').fillna(0)
    orders['phone_norm'] = orders.get('customer_mobile',pd.Series('',index=orders.index)).map(norm_phone)
    orders['order_date'] = pd.to_datetime(orders.get('added_at (Asia/Riyadh)'), errors='coerce')
    orders['method_label'] = orders.get('payment_method',pd.Series('',index=orders.index)).fillna('غير محدد').astype(str)
    orders['amount_num'] = pd.to_numeric(orders.get('total'),errors='coerce').fillna(0)
    orders['reference_clean'] = orders.get('transaction_reference',pd.Series('',index=orders.index)).fillna('').astype(str).str.strip()
    return raw, orders

def make_template():
    cols=['provider','transaction_id','order_reference','phone','customer_name','gross_amount','fee','net_amount','settlement_status','settlement_date','deposit_reference']
    return pd.DataFrame(columns=cols)

def reconcile(orders, payments):
    out=orders.copy()
    out['match_status']='غير موجود في ملف المنصة'
    out['matched_provider']=''
    out['provider_gross']=pd.NA; out['provider_fee']=pd.NA; out['provider_net']=pd.NA
    out['provider_settlement']=''; out['provider_settlement_date']=''; out['provider_reference']=''
    if payments is None or payments.empty: return out
    p=payments.copy()
    for c in ['provider','transaction_id','order_reference','phone','customer_name','settlement_status','settlement_date','deposit_reference']:
        if c not in p: p[c]=''
        p[c]=p[c].fillna('').astype(str).str.strip()
    for c in ['gross_amount','fee','net_amount']:
        if c not in p: p[c]=0
        p[c]=pd.to_numeric(p[c],errors='coerce').fillna(0)
    p['phone_norm']=p['phone'].map(norm_phone)
    p['order_reference_clean']=p['order_reference'].str.replace(r'\.0$','',regex=True)
    # Prefer exact reference; fallback to normalized mobile + amount tolerance.
    ref_map={}
    for _,r in p[p.order_reference_clean!=''].iterrows(): ref_map.setdefault(r.order_reference_clean,[]).append(r)
    for idx,r in out.iterrows():
        candidates=ref_map.get(str(r['id']),[])
        if not candidates and r['reference_clean']:
            candidates=ref_map.get(r['reference_clean'],[])
        if not candidates and r['phone_norm']:
            candidates=[x for _,x in p[p.phone_norm==r['phone_norm']].iterrows() if abs(float(x.gross_amount)-float(r.amount_num))<=0.02]
        if candidates:
            x=candidates[0]
            out.at[idx,'matched_provider']=x.provider
            out.at[idx,'provider_gross']=float(x.gross_amount)
            out.at[idx,'provider_fee']=float(x.fee)
            out.at[idx,'provider_net']=float(x.net_amount)
            out.at[idx,'provider_settlement']=x.settlement_status
            out.at[idx,'provider_settlement_date']=x.settlement_date
            out.at[idx,'provider_reference']=x.transaction_id or x.order_reference
            delta=abs(float(x.gross_amount)-float(r.amount_num))
            if delta>0.02: status='فرق في المبلغ'
            elif any(t in x.settlement_status.lower() for t in ['settled','paid','completed','تم','مكتمل','مسوى']): status='مطابق - تمت التسوية'
            elif x.settlement_status: status='مطابق - حالة التسوية تحتاج تحقق'
            else: status='مطابق - حالة التسوية غير محددة'
            out.at[idx,'match_status']=status
    return out

st.title('مطابقة الشويرد المالية')
st.caption('مطابقة طلبات زد مع ملفات منصات الدفع، مع تقرير قابل للتصدير.')
with st.sidebar:
    st.header('مصادر البيانات')
    zid_file=st.file_uploader('ارفع ملف زد (Excel)',type=['xlsx','xls'])
    st.divider()
    st.subheader('ملفات منصات الدفع')
    st.caption('يمكن رفع ملف موحّد بصيغة القالب أدناه. الربط المباشر بالـ API يضاف بعد التحقق من صلاحيات تقارير العمليات والتسويات لكل مزود.')
    pay_file=st.file_uploader('ارفع ملف عمليات الدفع (CSV أو Excel)',type=['csv','xlsx'])
    st.download_button('تحميل قالب عمليات الدفع',data=make_template().to_csv(index=False).encode('utf-8-sig'),file_name='payment_transactions_template.csv',mime='text/csv',use_container_width=True)

if not zid_file:
    st.info('ابدأ برفع ملف زد لعرض الطلبات وتجهيز المطابقة.')
    st.markdown('**المنصات المستهدفة:** '+ '، '.join(PROVIDERS))
    st.stop()

try:
    raw, orders=read_zid(zid_file)
except Exception as e:
    st.error(f'تعذرت قراءة الملف: {e}'); st.stop()

payments=None
if pay_file:
    try:
        payments=pd.read_csv(pay_file) if pay_file.name.lower().endswith('.csv') else pd.read_excel(pay_file)
        required={'provider','gross_amount'}
        if not required.issubset(set(payments.columns)):
            st.warning('ملف المنصة لا يطابق القالب. يلزم على الأقل عمودا provider و gross_amount.'); payments=None
    except Exception as e:
        st.error(f'تعذرت قراءة ملف الدفع: {e}')

result=reconcile(orders,payments)
# Suspicious duplicates: identical phone + amount + same normalized customer name, not automatic fraud assertion.
name_col='customer_name' if 'customer_name' in result else None
result['name_key']=result[name_col].fillna('').astype(str).str.strip().str.lower() if name_col else ''
result['duplicate_suspect']=False
mask=result['phone_norm'].ne('') & result['name_key'].ne('')
result.loc[mask,'duplicate_suspect']=result.loc[mask].duplicated(['phone_norm','amount_num','name_key'],keep=False)

min_date=result.order_date.min(); max_date=result.order_date.max()
with st.container():
    a,b,c,d=st.columns(4)
    a.metric('عدد الطلبات',f'{len(result):,}')
    b.metric('إجمالي الطلبات',f"{result.amount_num.sum():,.2f} ر.س")
    c.metric('مطابق مع المنصة',f"{result.match_status.str.startswith('مطابق').sum():,}")
    d.metric('يحتاج مراجعة',f"{((result.match_status!='مطابق - تمت التسوية') | result.duplicate_suspect).sum():,}")

st.subheader('تصفية النتائج')
f1,f2,f3,f4=st.columns(4)
with f1:
    if pd.notna(min_date) and pd.notna(max_date):
        date_range=st.date_input('تاريخ الطلب',value=(min_date.date(),max_date.date()))
    else: date_range=None
with f2: method=st.multiselect('طريقة الدفع',sorted(result.method_label.unique().tolist()))
with f3: statuses=st.multiselect('حالة المطابقة',sorted(result.match_status.unique().tolist()))
with f4: order_search=st.text_input('رقم الطلب')
filtered=result.copy()
if isinstance(date_range,tuple) and len(date_range)==2:
    filtered=filtered[filtered.order_date.dt.date.between(date_range[0],date_range[1])]
if method: filtered=filtered[filtered.method_label.isin(method)]
if statuses: filtered=filtered[filtered.match_status.isin(statuses)]
if order_search: filtered=filtered[filtered.id.astype(str).str.contains(order_search,na=False)]

st.subheader('نتائج المطابقة')
show_cols=['id','customer_mobile','payment_method','total','transaction_reference','order_date','matched_provider','provider_gross','provider_fee','provider_net','provider_settlement','provider_settlement_date','match_status','duplicate_suspect']
show_cols=[c for c in show_cols if c in filtered.columns]
st.dataframe(filtered[show_cols].rename(columns={
'id':'رقم الطلب','customer_mobile':'الجوال','payment_method':'طريقة الدفع','total':'مبلغ الطلب','transaction_reference':'مرجع زد','order_date':'تاريخ الطلب','matched_provider':'المنصة','provider_gross':'المبلغ بالمنصة','provider_fee':'الرسوم','provider_net':'الصافي','provider_settlement':'حالة التسوية','provider_settlement_date':'تاريخ التسوية','match_status':'نتيجة المطابقة','duplicate_suspect':'تكرار محتمل'}),use_container_width=True,hide_index=True)

st.subheader('تنبيهات المراجعة')
dups=result[result.duplicate_suspect]
if len(dups): st.warning(f'يوجد {len(dups)} طلباً ضمن مجموعات متشابهة (نفس الجوال والمبلغ والاسم). يلزم مراجعتها ولا تعتبر تكراراً مؤكداً.')
else: st.success('لا توجد مجموعات تكرار واضحة حسب الجوال والمبلغ والاسم في الملف.')

buf=io.BytesIO()
with pd.ExcelWriter(buf,engine='openpyxl') as writer:
    filtered.drop(columns=['name_key'],errors='ignore').to_excel(writer,index=False,sheet_name='نتائج المطابقة')
    result[result.duplicate_suspect].drop(columns=['name_key'],errors='ignore').to_excel(writer,index=False,sheet_name='تكرار محتمل')
    pd.DataFrame({'المؤشر':['عدد الطلبات','إجمالي الطلبات','طلبات مطابقة','طلبات غير مطابقة','تكرار محتمل'], 'القيمة':[len(result),round(result.amount_num.sum(),2),int(result.match_status.str.startswith('مطابق').sum()),int((result.match_status=='غير موجود في ملف المنصة').sum()),int(result.duplicate_suspect.sum())]}).to_excel(writer,index=False,sheet_name='ملخص')
st.download_button('تحميل تقرير Excel',data=buf.getvalue(),file_name='alshwered_reconciliation.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')

with st.expander('حالة الربط المباشر بالمنصات'):
    st.write('تم تجهيز بنية المطابقة ورفع ملفات العمليات. لا يتم إرسال طلبات API تلقائياً قبل تحديد واجهات التقارير المعتمدة وصلاحيات حساب التاجر؛ مفاتيح التهيئة وحدها لا تثبت أن صلاحية التسويات متاحة.')
    st.write('مراجع التهيئة: Tamara API docs وZid payment provider setup. راجع الإعدادات الرسمية قبل تفعيل الإنتاج.')
