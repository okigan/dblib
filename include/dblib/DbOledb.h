#if !defined(AFX_COLEDB_H__20011208_FD4E_9BC2_32F3_0080AD509054__INCLUDED_)
#define AFX_COLEDB_H__20011208_FD4E_9BC2_32F3_0080AD509054__INCLUDED_

/////////////////////////////////////////////////////////////////////////////
// DbOledb.h - OLEDB wrapping classes
//
// Written by Bjarke Viksoe (bjarke@viksoe.dk)
// Copyright (c) 2001 Bjarke Viksoe.
//
// Requires inclusion of the DbBase.h header.
//
// This code may be used in compiled form in any way you desire. This
// file may be redistributed by any means PROVIDING it is 
// not sold for profit without the authors written consent, and 
// providing that this notice and the authors name is included. 
//
// This file is provided "as is" with no expressed or implied warranty.
// The author accepts no liability if it causes any damage to you or your
// computer whatsoever. It's free, so don't hassle me about it.
//
// Beware of bugs.
//

#pragma once

#ifndef __ATLBASE_H__
#error OLEDB database classes requires inclusion of atlbase.h
#endif

#include <oledb.h>         // OLE DB Header
#include <oledberr.h>      // OLE DB Errors
#include <msdasc.h>        // OLE DB Service Component header
#include <msdaguid.h>      // OLE DB Root Enumerator
#include <msdasql.h>       // MSDASQL - Default provider
#pragma comment(lib, "oledb.lib")
//#pragma comment(lib, "oledbd.lib")

#include "DbBase.h"


class COledbDatabase;
class COledbRecordset;
class COledbCommand;
class COledbErrors;
class COledbError;


class COledbSystem : 
    public IDbSystem
  , public CComObjectRoot
  , public CComCoClass<COledbSystem>
{
public:
   COledbSystem();
   virtual ~COledbSystem();

   BEGIN_COM_MAP(COledbSystem)
   END_COM_MAP()


   BOOL Initialize();
   void Terminate();

   IDbDatabase* CreateDatabase();
   IDbRecordset* CreateRecordset(IDbDatabase* pDb);
   IDbCommand* CreateCommand(IDbDatabase* pDb);
};


class COledbError : public IDbError
{
friend COledbErrors;
protected:
   CComBSTR m_bstrSource;
   CComBSTR m_bstrMsg;
   long m_lNative;

public:
   COledbError();
   virtual ~COledbError();

   long GetErrorCode();
   long GetNativeErrorCode();
   void GetOrigin(LPTSTR pstrStr, UINT cchMax);
   void GetSource(LPTSTR pstrStr, UINT cchMax);
   void GetMessage(LPTSTR pstrStr, UINT cchMax);

protected:
   void _Init(long lNative, BSTR bstrSource, BSTR bstrMsg);
};


class COledbErrors : public IDbErrors
{
friend COledbDatabase;
protected:
   enum { MAX_ERRORS = 5 };
   COledbError m_p[MAX_ERRORS];
   long m_lCount;

public:
   COledbErrors();
   virtual ~COledbErrors();

   void Clear();
   long GetCount() const;
   IDbError* GetError(short iIndex);

protected:
   void _Init(HRESULT Hr);
};


class COledbDatabase : 
    public IDbDatabase
  , public CComObjectRoot
  , public CComCoClass<COledbDatabase>
{
friend COledbRecordset;
friend COledbCommand;
public:
   CComPtr<IDBInitialize> m_spInit;
   CComPtr<IOpenRowset> m_spSession;
protected:
   COledbSystem* m_pSystem;
   COledbErrors* m_pErrors;
   long m_lLoginTimeout;
   long m_lQueryTimeout;
#ifdef _DEBUG
   long m_nRecordsets;
#endif

public:
   COledbDatabase();
   virtual ~COledbDatabase();

   void SetSystem(COledbSystem* pSystem);

   BEGIN_COM_MAP(COledbSystem)
   END_COM_MAP()

   //IDbDatabase
   BOOL Open(HWND hWnd, LPCTSTR pstrConnectionString, LPCTSTR pstrUser, LPCTSTR pstrPassword, long iType = DB_OPEN_DEFAULT);
   void Close();
   BOOL IsOpen() const;
   BOOL ExecuteSQL(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT, DWORD* pdwRowsAffected = NULL);
   BOOL BeginTrans();
   BOOL CommitTrans();
   BOOL RollbackTrans();
   void SetLoginTimeout(long lTimeout);
   void SetQueryTimeout(long lTimeout);
   IDbErrors* GetErrors();

public:
   BOOL BeginTrans(ISOLEVEL isoLevel);
   BOOL Connect();

protected:
   BOOL _Error(HRESULT Hr);
   static void _SetRecordsetType(ICommand* pCommand, long lType, long lOptions);
};


class COledbRecordset : 
    public IDbRecordset
  , public CComObjectRoot
  , public CComCoClass<COledbRecordset>
{
friend COledbCommand;
public:
   CComPtr<IRowset> m_spRowset;
   DBBINDING* m_rgBindings;
   HACCESSOR m_hAccessor;
   LPVOID m_pData;
   DWORD m_dwBufferSize;
   LPWSTR m_pwstrNameBuffer;
protected:
   COledbDatabase* m_pDb;
   LONG m_nRowsAffected;
   short m_nCols;
   short m_iAdjustIndex;
   bool m_fEOF;
   long m_lType;
   long m_lOptions;
   long m_lQueryTimeout;

public:
   COledbRecordset(IRowset* pRS);
   COledbRecordset();
   virtual ~COledbRecordset();

   void SetDatabase(COledbDatabase* pDb);

   BEGIN_COM_MAP(COledbSystem)
   END_COM_MAP()

   //IDbRecordset
   BOOL Open(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT);
   void Close();
   BOOL IsOpen() const;

   DWORD GetRowCount() const;
   BOOL GetField(short iIndex, long& pData);
   BOOL GetField(short iIndex, float& pData);
   BOOL GetField(short iIndex, double& pData);
   BOOL GetField(short iIndex, bool& pData);
   BOOL GetField(short iIndex, LPTSTR pData, UINT cchMax);
   BOOL GetField(short iIndex, SYSTEMTIME& pSt);
   BOOL GetColumnName(short iIndex, LPTSTR pstrName, UINT cchMax);
   short GetColumnType(short iIndex);
   short GetColumnIndex(LPCTSTR pstrName) const;
   DWORD GetColumnSize(short iIndex);
   DWORD GetColumnCount() const;
   BOOL IsEOF() const;
   BOOL MoveNext();
   BOOL MovePrev();
   BOOL MoveTop();
   BOOL MoveBottom();
   BOOL MoveAbs(DWORD dwPos);
   DWORD GetRowNumber();
   BOOL NextResultset();

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)
   BOOL GetField(short iIndex, CString& pData);
#endif // __ATLSTR_H__
#if defined(_STRING_)
   BOOL GetField(short iIndex, std::string& pData);
#endif // __ATLSTR_H__

public:
   BOOL Attach(IRowset* pRowset);
   BOOL MoveCursor(LONG lSkip, LONG lAmount);

protected:
   BOOL _Error(HRESULT Hr);
   BOOL _BindColumns();
   BOOL _Attach(IRowset* pRowset);
};

class COledbCommand : public IDbCommand
{
public:
   CComPtr<IRowset> m_spRowset;
   CComQIPtr<ICommandText> m_spText;
protected:
   enum { MAX_PARAMS = 20 };
   enum { MAX_PARAMBUFFER_SIZE = 1024 };
   COledbDatabase* m_pDb;
   LONG m_nRowsAffected;
   short m_nParams;
   DBBINDING* m_rgBindings;
   HACCESSOR m_hAccessor;
   LPVOID m_pData;
   DWORD m_dwBindOffset;

public:
   COledbCommand(COledbDatabase* pDb);
   virtual ~COledbCommand();

   BOOL Create(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT);
   BOOL Execute(IDbRecordset* pRecordset = NULL);
   BOOL Cancel();
   void Close();
   BOOL IsOpen() const;
   DWORD GetRowCount() const;
   BOOL SetParam(short iIndex, const long* pData);
   BOOL SetParam(short iIndex, LPCTSTR pData, UINT cchMax = -1);
   BOOL SetParam(short iIndex, const bool* pData);
   BOOL SetParam(short iIndex, const float* pData);
   BOOL SetParam(short iIndex, const double* pData);
   BOOL SetParam(short iIndex, const SYSTEMTIME* pData, short iType = DB_TYPE_TIMESTAMP);
#if defined(_STRING_)
   BOOL SetParam(short iIndex, std::string& str);
#endif // _STRING
#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)
   BOOL SetParam(short iIndex, CString& str);
#endif // __ATLSTR_H__

protected:
   BOOL _Error(HRESULT Hr);
};


#endif // !defined(AFX_COLEDB_H__20011208_FD4E_9BC2_32F3_0080AD509054__INCLUDED_)
