#pragma once
#import "C:\Program Files\Common Files\System\ado\msado15.dll"  no_namespace rename("EOF","adoEOF") rename("BOF","adoBOF")
class CADO
{
public:
	CADO();
	virtual ~CADO();
private:
	_ConnectionPtr m_pConn;
	_RecordsetPtr  m_pRS;
public:
	_RecordsetPtr& GetRS(CString strSQL);	// 执行  返回记录集的SQL语句
	BOOL Execute(CString strSQL);			// 执行不返回记录集的SQL语句

};