

CADO ado;

_RecordsetPtr pSet_gdt;
CString str_gdt;
str_gdt.Format(_T("Select gdt_test.gtno,gdt_standard.gname,gdt_test.pre,gdt_test.pw,gdt_standard.gdcsv,gdt_test.isv,gdt_test.rt,gdt_test.av from gdt_test,gdt_standard where gdt_standard.gdcsv >= " + _bstr_t(wv_gdt) + "and gdt_test.isv <= " + _bstr_t(m_mv) + "and gdt_test.av <=" + _bstr_t(m_cv)));
pSet_gdt = ado.GetRS(str_gdt);
while (!pSet_gdt->adoEOF)
{

	m_ListControl_gdt.InsertItem(0, _T(" "));
	for (int i = 0; i <= 7; ++i)
	{
		m_ListControl_gdt.SetItemText(0, i, (_bstr_t)pSet_gdt->GetCollect(_variant_t((long)i)));
	}
	pSet_gdt->MoveNext();
}