#if !defined(AFX_CSqlite3_H__20040122_8927_6EF3_B52C_0080AD509054__INCLUDED_)
#define AFX_CSqlite3_H__20040122_8927_6EF3_B52C_0080AD509054__INCLUDED_

/////////////////////////////////////////////////////////////////////////////
// DbSqlite3.h - Sqlite v3 wrapping classes
//
// Written by Bjarke Viksoe (bjarke@viksoe.dk)
// Copyright (c) 2005 Bjarke Viksoe.
//
// BUG: This code ignores the fact that the 8-bit character interface
//      to Sqlite expects UTF-8 encoding. The database itself doesn't
//      fail if it gets ASCII/multibyte strings, but it may not
//      always function as expected (sorting, string compares, etc).
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

#include "sqlite3.h"

// This SQLite3 wrappers default to using the DLL/.lib file (create it
// with "lib /def:sqlite3.def"). Alternatively you can compile
// a static lib and include it with the _USE_SQLITE3_LIBS define.
#ifdef _USE_SQLITE3_LIBS
   #ifdef _DEBUG
      #pragma comment(lib, "sqlite3_mtd.lib")
   #else
      #pragma comment(lib, "sqlite3_mtr.lib")
   #endif
#else
   #pragma comment(lib, "sqlite3.lib")
#endif

#include "DbBase.h"

#ifndef DB_MAX_PARAMS
   #define DB_MAX_PARAMS 25
#endif // DB_MAX_PARAMS


// Forware declares
class CSqlite3Database;
class CSqlite3Recordset;
class CSqlite3Command;
class CSqlite3Errors;
class CSqlite3Error;


class CSqlite3System : public IDbSystem
{
public:
   BOOL Initialize();
   void Terminate();

   IDbDatabase*  CreateDatabase();
   IDbRecordset* CreateRecordset(IDbDatabase* pDb);
   IDbCommand*   CreateCommand(IDbDatabase* pDb);
};


class CSqlite3Error : public IDbError
{
friend CSqlite3Errors;
protected:
   int m_iError;
   int m_iNative;
   CHAR m_szMsg[300];

public:
   long GetErrorCode();
   long GetNativeErrorCode();
   void GetOrigin(LPTSTR pstrStr, UINT cchMax);
   void GetSource(LPTSTR pstrStr, UINT cchMax);
   void GetMessage(LPTSTR pstrStr, UINT cchMax);

protected:
   void _Init(int iError, int iNative, LPCSTR pstrMessage);
};


class CSqlite3Errors : public IDbErrors
{
friend CSqlite3Database;
protected:
   CSqlite3Error m_Error;
   long m_lCount;

public:
   void Clear();
   long GetCount() const;
   IDbError* GetError(short iIndex);

protected:
   void _Init(int iError, int iNative, LPCSTR pstrMessage);
};


class CSqlite3Database : public IDbDatabase
{
friend CSqlite3Recordset;
friend CSqlite3Command;
public:
   sqlite3* m_pDb;
protected:
   CSqlite3System* m_pSystem;
   CSqlite3Errors m_Errors;
#ifdef _DEBUG
   long m_nRecordsets;
#endif

public:
   CSqlite3Database(CSqlite3System* pSystem);
   virtual ~CSqlite3Database();

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

public:
   INT64 GetLastRowId();
   operator struct sqlite3*() { return m_pDb; };

protected:
   BOOL _Error(int iError, int iNative, LPCSTR pstrMessage);
};


class CSqlite3Recordset : public IDbRecordset
{
friend CSqlite3Command;
public:
protected:
   CSqlite3Database* m_pDb;
   sqlite3_stmt* m_pVm;
   CHAR** m_ppSnapshot;
   const CHAR** m_ppColumns;
   int m_nCols;
   int m_nRows;
   int m_iPos;
   long m_lType;
   long m_lOptions;
   BOOL m_fEOF;
   bool m_bAttached;

public:
   CSqlite3Recordset(CSqlite3Database* pDb);
   virtual ~CSqlite3Recordset();

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

protected:
   BOOL _Error(int iError, int iNative, LPCSTR pstrMessage);
   BOOL _Attach(sqlite3_stmt* pVm);
};


class CSqlite3Command : public IDbCommand
{
protected:
   struct TSqlite3Param
   {
      VARTYPE vt;
      LPCVOID pVoid;
      int len;
   };
protected:
   CSqlite3Database* m_pDb;
   sqlite3_stmt* m_pVm;
   LPTSTR m_pszSQL;
   short m_nParams;
   TSqlite3Param m_Params[DB_MAX_PARAMS];

public:
   CSqlite3Command(CSqlite3Database* pDb);
   virtual ~CSqlite3Command();

// IDbCommand methods
public:
   BOOL Create(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT);
   BOOL Execute(IDbRecordset* pRecordset = NULL);
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

protected:
   BOOL _Error(int iError, int iNative, LPCSTR pstrMessage);
};


#endif // !defined(AFX_CSqlite3_H__20040122_8927_6EF3_B52C_0080AD509054__INCLUDED_)
