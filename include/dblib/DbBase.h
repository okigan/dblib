#if !defined(AFX_IDb_H__20011208_E6E3_45BF_60C8_0080AD509054__INCLUDED_)
#define AFX_IDb_H__20011208_E6E3_45BF_60C8_0080AD509054__INCLUDED_

/////////////////////////////////////////////////////////////////////////////
// DbBase.h - Generic database interfaces
//
// Written by Bjarke Viksoe (bjarke@viksoe.dk)
// Copyright (c) 2001-2005 Bjarke Viksoe.
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


// Database open options
#define DB_OPEN_DEFAULT              0
#define DB_OPEN_READ_ONLY            1
#define DB_OPEN_SILENT               4
#define DB_OPEN_PROMPTDIALOG         8
#define DB_OPEN_USECURSORS          16

// Recordset open types
#define DB_OPEN_TYPE_DYNASET         1
#define DB_OPEN_TYPE_SNAPSHOT        2
#define DB_OPEN_TYPE_FORWARD_ONLY    3
#define DB_OPEN_TYPE_UPDATEONLY      4

// Recordset open options
#define DB_OPTION_DEFAULT            0
#define DB_OPTION_APPEND_ONLY        1
#define DB_OPTION_PREPARE            2

// Database column types
#define DB_TYPE_UNKNOWN              0
#define DB_TYPE_INTEGER              1
#define DB_TYPE_REAL                 2
#define DB_TYPE_DOUBLE               3
#define DB_TYPE_CHAR                 4
#define DB_TYPE_DATE                 5
#define DB_TYPE_TIME                 6
#define DB_TYPE_TIMESTAMP            7
#define DB_TYPE_BOOLEAN              8


#ifndef _ASSERTE
#define _ASSERTE(x)
#endif // _ASSERTE

#ifndef ATLTRY
#define ATLTRY(x) (x)
#endif // ATLTRY

class IDbDatabase;
class IDbRecordset;
class IDbCommand;

class IDbSystem
{
public:
   virtual BOOL Initialize() = 0;
   virtual void Terminate() = 0;
   virtual IDbDatabase*  CreateDatabase() = 0;
   virtual IDbRecordset* CreateRecordset(IDbDatabase* pDb) = 0;
   virtual IDbCommand*   CreateCommand(IDbDatabase* pDb) = 0;
};

class IDbError
{
public:
   virtual long GetErrorCode() = 0;
   virtual long GetNativeErrorCode() = 0;
   virtual void GetOrigin(LPTSTR pstrStr, UINT cchMax) = 0;
   virtual void GetSource(LPTSTR pstrStr, UINT cchMax) = 0;
   virtual void GetMessage(LPTSTR pstrStr, UINT cchMax) = 0;
};

class IDbErrors
{
public:
   virtual void Clear() = 0;
   virtual long GetCount() const = 0;
   virtual IDbError* GetError(short iIndex) = 0;
};

class IDbDatabase
{
public:
   virtual BOOL Open(HWND hWnd, LPCTSTR pstrConnectionString, LPCTSTR pstrUser, LPCTSTR pstrPassword, long iType = DB_OPEN_DEFAULT) = 0;
   virtual void Close() = 0;
   virtual BOOL ExecuteSQL(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT, DWORD* pdwRowsAffected = NULL) = 0;
   virtual BOOL BeginTrans() = 0;
   virtual BOOL CommitTrans() = 0;
   virtual BOOL RollbackTrans() = 0;
   virtual BOOL IsOpen() const = 0;
   virtual IDbErrors* GetErrors() = 0;
};

class IDbRecordset
{
public:
   virtual BOOL Open(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT) = 0;
   virtual void Close() = 0;
   virtual BOOL IsOpen() const = 0;
   virtual DWORD GetRowCount() const = 0;
   virtual BOOL GetField(short iIndex, long& Data) = 0;
   virtual BOOL GetField(short iIndex, float& Data) = 0;
   virtual BOOL GetField(short iIndex, double &Data) = 0;
   virtual BOOL GetField(short iIndex, bool &Data) = 0;
   virtual BOOL GetField(short iIndex, LPTSTR pData, UINT cchMax) = 0;
   virtual BOOL GetField(short iIndex, SYSTEMTIME& Data) = 0;
   virtual DWORD GetColumnSize(short iIndex) = 0;
   virtual BOOL GetColumnName(short iIndex, LPTSTR pstrName, UINT cchMax) = 0;
   virtual short GetColumnType(short iIndex) = 0;
   virtual short GetColumnIndex(LPCTSTR pstrName) const = 0;
   virtual DWORD GetColumnCount() const = 0;
   virtual BOOL IsEOF() const = 0;
   virtual BOOL MoveNext() = 0;
   virtual BOOL MovePrev() = 0;
   virtual BOOL MoveAbs(DWORD dwPos) = 0;
   virtual BOOL MoveTop() = 0;
   virtual BOOL MoveBottom() = 0;
   virtual DWORD GetRowNumber() = 0;
   virtual BOOL NextResultset() = 0;

//#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)
//   virtual BOOL GetField(short iIndex, CString& pData) = 0;
//#endif // __ATLSTR_H__
//#if defined(_STRING_)
//   virtual BOOL GetField(short iIndex, std::string& pData) = 0;
//#endif // __ATLSTR_H__

};

class IDbCommand
{
public:
   virtual BOOL Create(LPCTSTR pstrSQL, long lType = DB_OPEN_TYPE_FORWARD_ONLY, long lOptions = DB_OPTION_DEFAULT) = 0;
   virtual BOOL Execute(IDbRecordset* pRecordset = NULL) = 0;
   virtual void Close() = 0;
   virtual BOOL IsOpen() const = 0;
   virtual DWORD GetRowCount() const = 0;
   virtual BOOL SetParam(short iIndex, const long* pData) = 0;
   virtual BOOL SetParam(short iIndex, LPCTSTR pData, UINT cchMax = -1) = 0;
   virtual BOOL SetParam(short iIndex, const bool* pData) = 0;
   virtual BOOL SetParam(short iIndex, const float* pData) = 0;
   virtual BOOL SetParam(short iIndex, const double* pData) = 0;
   virtual BOOL SetParam(short iIndex, const SYSTEMTIME* pData, short iType = DB_TYPE_TIMESTAMP) = 0;
#if 0
#if defined(_STRING_)
   virtual BOOL SetParam(short iIndex, std::string& str) = 0;
#endif // _STRING
#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)
   virtual BOOL SetParam(short iIndex, CString& str) = 0;
#endif // __ATLSTR_H__
#endif
};


#ifndef _NO_DB_HELPERS

class IDbAutoTrans
{
public:
   IDbDatabase* m_pDb;
public:
   IDbAutoTrans(IDbDatabase* pDb) : m_pDb(pDb)
   {
      _ASSERTE(m_pDb);
      _ASSERTE(m_pDb->IsOpen());
      m_pDb->BeginTrans();
   }
   ~IDbAutoTrans()
   {
      // Automatically call Rollback(). You must
      // have called Commit() at some point to commit the transaction.
      Rollback();
   }
   BOOL Commit()
   {
      if( m_pDb == NULL ) return FALSE;
      _ASSERTE(m_pDb->IsOpen());
      BOOL bRes = m_pDb->CommitTrans();
      m_pDb = NULL;
      return bRes;
   }
   BOOL Rollback()
   {
      if( m_pDb == NULL ) return FALSE;
      _ASSERTE(m_pDb->IsOpen());
      BOOL bRes = m_pDb->RollbackTrans();
      m_pDb = NULL;
      return bRes;
   }
};

#endif // _NO_DB_HELPERS


#endif // !defined(AFX_IDb_H__20011208_E6E3_45BF_60C8_0080AD509054__INCLUDED_)
