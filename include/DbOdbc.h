#if !defined(AFX_CODBC_H__20011208_ED4E_9BC2_32F3_0080AD509054__INCLUDED_)
#define AFX_CODBC_H__20011208_ED4E_9BC2_32F3_0080AD509054__INCLUDED_

/////////////////////////////////////////////////////////////////////////////
// DbOdbc.h - ODBC wrapping classes
//
// Written by Bjarke Viksoe (bjarke@viksoe.dk)
// Copyright (c) 2001-2003 Bjarke Viksoe.
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

#include <sql.h>
#include <sqlext.h>
#include <odbcinst.h>
#pragma comment(lib, "odbc32.lib")

#include "DbBase.h"

#ifndef DB_MAX_PARAMS
   #define DB_MAX_PARAMS 25
#endif // DB_MAX_PARAMS


class COdbcDatabase;
class COdbcRecordset;
class COdbcCommand;
class COdbcErrors;
class COdbcError;


class COdbcSystem : public IDbSystem
{
public:
   HENV m_henv;

public:
   COdbcSystem();
   virtual ~COdbcSystem();

   BOOL Initialize();
   void Terminate();

   IDbDatabase*  CreateDatabase();
   IDbRecordset* CreateRecordset(IDbDatabase* pDb);
   IDbCommand*   CreateCommand(IDbDatabase* pDb);

// ODBC specific methods
public:
   BOOL SetConnectionPooling();
};


class COdbcError : public IDbError
{
friend COdbcErrors;
protected:
   TCHAR m_szMsg[SQL_MAX_MESSAGE_LENGTH+1];
   TCHAR m_szState[SQL_SQLSTATE_SIZE+1];
   long m_lNative;

public:
   COdbcError();
   virtual ~COdbcError();

   long GetErrorCode();
   long GetNativeErrorCode();
   void GetOrigin(LPTSTR pstrStr, UINT cchMax);
   void GetSource(LPTSTR pstrStr, UINT cchMax);
   void GetMessage(LPTSTR pstrStr, UINT cchMax);

// ODBC specific methods
public:
   void GetState(LPTSTR pstrStr, UINT cchMax);

protected:
   void _Init(LPCTSTR pstrState, long lNative, LPCTSTR pstrMsg);
};


class COdbcErrors : public IDbErrors
{
friend COdbcDatabase;
protected:
   enum { MAX_ERRORS = 6 };
   COdbcError m_p[MAX_ERRORS];
   long m_lCount;

public:
   COdbcErrors();
   virtual ~COdbcErrors();

   void Clear();
   long GetCount() const;
   IDbError* GetError(short iIndex);

protected:
   void _Init(HENV henv, HDBC hdbc, HSTMT hstmt);
};


class COdbcDatabase : public IDbDatabase
{
friend COdbcRecordset;
friend COdbcCommand;
public:
   HDBC m_hdbc;
   HSTMT m_hstmt;
protected:
   COdbcSystem* m_pSystem;
   COdbcErrors* m_pErrors;
   long m_lLoginTimeout;
   long m_lQueryTimeout;
#ifdef _DEBUG
   long m_nRecordsets;
#endif

public:
   COdbcDatabase(COdbcSystem* pSystem);
   virtual ~COdbcDatabase();

// IDbDatabase methods
public:
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

// ODBC specific methods
public:
   BOOL TranslateSQL(LPCTSTR pstrNativeSQL, LPTSTR pstrSQL, UINT cchMax);
   BOOL SetOption(SQLUSMALLINT Option, SQLULEN Value);
   BOOL SetTransactionLevel(SQLULEN Value);
   SQLULEN GetTransactionLevel();

protected:
   BOOL _Error(HSTMT hstmt=SQL_NULL_HSTMT);
   static void _SetRecordsetType(HSTMT hstmt, long lType, long lOptions);
};

class COdbcVariant
{
public:
   SHORT m_iType;
   SHORT m_iSqlType;
   DWORD m_dwSize;
   union {
      long lVal;
      float fVal;
      double dblVal;
      BOOL bVal;
      char* pstrVal;
      WCHAR* pwstrVal;
      DATE_STRUCT dateVal;
      TIME_STRUCT timeVal;
      TIMESTAMP_STRUCT stampVal;
   } data;

public:
   COdbcVariant();
   ~COdbcVariant();

   void Init(SHORT lType, DWORD lSize);
   void Clear();
   inline void Destroy();
};

class COdbcField
{
public:
   TCHAR m_szName[64];
   SQLSMALLINT m_iType;
   SQLSMALLINT m_iSqlType;
   SQLINTEGER m_dwSize;
   SQLLEN m_dwRetrivedSize;
   COdbcVariant m_val;

public:
   COdbcField(LPCTSTR pstrColName, SQLSMALLINT iType, SQLSMALLINT iSqlType, SQLINTEGER dwSize);
   ~COdbcField();

   LPVOID GetData() const;
};

class COdbcRecordset : public IDbRecordset
{
friend COdbcCommand;
public:
   HSTMT m_hstmt;
protected:
   COdbcDatabase* m_pDb;
   COdbcField** m_ppFields;
   short m_nCols;
   bool m_fEOF;
   bool m_fAttached;
   long m_lType;
   long m_lOptions;
   long m_lQueryTimeout;

public:
   COdbcRecordset(COdbcDatabase* pDb);
   virtual ~COdbcRecordset();

// IDbRecordset methods
public:
   BOOL Open(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT);
   void Close();
   BOOL IsOpen() const;

   DWORD GetRowCount() const;
   BOOL GetField(short iIndex, long& Data);
   BOOL GetField(short iIndex, float& Data);
   BOOL GetField(short iIndex, double &Data);
   BOOL GetField(short iIndex, bool &Data);
   BOOL GetField(short iIndex, LPTSTR pData, UINT cchMax);
   BOOL GetField(short iIndex, SYSTEMTIME& pSt);
#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)
   BOOL GetField(short iIndex, CString& pData);
#endif // __ATLSTR_H__
#if defined(_STRING_)
   BOOL GetField(short iIndex, std::string& pData);
#endif // __ATLSTR_H__
   DWORD GetColumnSize(short iIndex);
   BOOL GetColumnName(short iIndex, LPTSTR pstrName, UINT cchMax);
   short GetColumnType(short iIndex);
   short GetColumnIndex(LPCTSTR pstrName) const;
   DWORD GetColumnCount() const;
   BOOL IsEOF() const;
   BOOL MoveNext();
   BOOL MovePrev();
   BOOL MoveTop();
   BOOL MoveBottom();
   BOOL MoveAbs(DWORD dwPos);
   DWORD GetRowNumber();
   BOOL NextResultset();

// ODBC specific methods
public:
   BOOL Cancel();
   BOOL MoveCursor(SHORT lType, DWORD dwPos);
   BOOL GetCursorName(LPTSTR pszName, UINT cchMax) const;
   BOOL SetCursorName(LPCTSTR pszName);

protected:
   BOOL _Attach(HSTMT hStmt);
   BOOL _Error();
   BOOL _BindColumns();
   void _ConvertType(SQLSMALLINT& type, SQLUINTEGER& len) const;
};

typedef struct tagOdbcCommandValue
{
   LPCVOID pValue;
   short type;
   long size;
   // The union below is used in the conversion of the date
   // types to SYSTEMTIME type. A lot of work...
   union {
      TIMESTAMP_STRUCT ts;
      DATE_STRUCT date;
      TIME_STRUCT time;
   } dateVal;
} tOdbcCommandValue;

class COdbcCommand : public IDbCommand
{
public:
   HSTMT m_hstmt;
protected:
   COdbcDatabase* m_pDb;
   LPTSTR m_pszSQL;
   tOdbcCommandValue m_params[DB_MAX_PARAMS];
   long m_lOptions;
   long m_lQueryTimeout;
   short m_nParams;
   bool m_bAttached;

public:
   COdbcCommand(COdbcDatabase* pDb);
   virtual ~COdbcCommand();

// IDbCommand methods
public:
   BOOL Create(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT);
   BOOL Execute(IDbRecordset* pRecordset=NULL);
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
#endif _STRING
#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)
   BOOL SetParam(short iIndex, CString& str);
#endif // __ATLSTR_H__

// ODBC specific methods
public:
   BOOL SetOption(SQLUSMALLINT Option, SQLULEN Value);
   BOOL SetCursorName(LPCTSTR pszName);
   void SetQueryTimeout(long lTimeout);

protected:
   BOOL _Error();
};


#ifndef _NO_DB_HELPERS

class IDbAutoIsolationLevel
{
public:
   COdbcDatabase* m_pDb;
   SQLULEN m_iOldLevel;
public:
   IDbAutoIsolationLevel(COdbcDatabase* pDb, SQLULEN iNewLevel) : m_pDb(pDb)
   {
      _ASSERTE(m_pDb);
      _ASSERTE(m_pDb->IsOpen());
      m_iOldLevel = m_pDb->GetTransactionLevel();
      BOOL bRes = m_pDb->SetTransactionLevel(iNewLevel);
      _ASSERTE(bRes); bRes;
   }
   ~IDbAutoIsolationLevel()
   {
      BOOL bRes = m_pDb->SetTransactionLevel(m_iOldLevel);
      _ASSERTE(bRes); bRes;
   }
};

#endif // _NO_DB_HELPERS


#endif // !defined(AFX_CODBC_H__20011208_ED4E_9BC2_32F3_0080AD509054__INCLUDED_)
