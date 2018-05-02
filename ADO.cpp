#include "stdafx.h"
#include "ADO.h"


CADO::CADO()//str:�����ַ���
{
	::CoInitialize(NULL);				// ��ʼ��OLE/COM�⻷�� 
	try
	{
		m_pConn.CreateInstance("ADODB.Connection");	//����Connection����
		m_pConn->ConnectionTimeout = 5;	//���ó�ʱʱ��Ϊ5��
		m_pConn->Open("Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott;Data Source=orcl", "", "", 0);//�������ݿ�
	}
	catch (_com_error e)					//��׽�쳣
	{  
		CString errormessage;
		errormessage.Format(_T("�������ݿ�ʧ��!\n������Ϣ:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);
	}
}
/*********************************************************************/
CADO::~CADO()
{
	if (m_pRS != NULL)				// �رռ�¼��������
		m_pRS->Close();
	m_pConn->Close();

	::CoUninitialize();				// �ͷŻ���
}

/*********************************************************************/
_RecordsetPtr&  CADO::GetRS(CString strSQL) //ִ��strSQL��SQL��䣬���ؼ�¼��
{
	try
	{
		m_pRS.CreateInstance(__uuidof(Recordset));
		m_pRS->Open((_bstr_t)strSQL, m_pConn.GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdText);	// ȡ�ñ��еļ�¼
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(_T("���ؼ�¼��ʧ��!\n������Ϣ:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);
	}
	return m_pRS;
}
/*********************************************************************/

BOOL CADO::Execute(CString strSQL)//ִ��strSQL��SQL��䣬�����ؼ�¼��
{
	try
	{
		m_pConn->Execute((_bstr_t)strSQL, NULL, adCmdText);
		return true;
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(_T("ִ��SQLʧ��!\n������Ϣ:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);
		return false;
	}
}
