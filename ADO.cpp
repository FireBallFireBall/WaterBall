#include "stdafx.h"
#include "ADO.h"


CADO::CADO()//str:连接字符串
{
	::CoInitialize(NULL);				// 初始化OLE/COM库环境 
	try
	{
		m_pConn.CreateInstance("ADODB.Connection");	//创建Connection对象
		m_pConn->ConnectionTimeout = 5;	//设置超时时间为5秒
		m_pConn->Open("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl", "", "", 0);//连接数据库
	}
	catch (_com_error e)					//捕捉异常
	{  
		CString errormessage;
		errormessage.Format(_T("连接数据库失败!\n错误信息:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);
	}
}
/*********************************************************************/
CADO::~CADO()
{
	if (m_pRS != NULL)				// 关闭记录集和连接
		m_pRS->Close();
	m_pConn->Close();

	::CoUninitialize();				// 释放环境
}

/*********************************************************************/
_RecordsetPtr&  CADO::GetRS(CString strSQL) //执行strSQL的SQL语句，返回集录集
{
	try
	{
		m_pRS.CreateInstance(__uuidof(Recordset));
		m_pRS->Open((_bstr_t)strSQL, m_pConn.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);	// 取得表中的记录
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(_T("返回集录集失败!\n错误信息:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);
	}
	return m_pRS;
}
/*********************************************************************/

BOOL CADO::Execute(CString strSQL)//执行strSQL的SQL语句，不返回集录集
{
	try
	{
		m_pConn->Execute((_bstr_t)strSQL, NULL, adCmdText);
		return true;
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(_T("执行SQL失败!\n错误信息:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);
		return false;
	}
}
