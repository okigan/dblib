#if !defined(AFX_CCSV_H__20011208_ED4E_9BC2_32F3_0080AD509054__INCLUDED_)
#define AFX_CCSV_H__20011208_ED4E_9BC2_32F3_0080AD509054__INCLUDED_

/////////////////////////////////////////////////////////////////////////////
// cCsv.h - CSV wrapping classes
//
// Written by Bjarke Viksoe (bjarke@viksoe.dk)
// Copyright (c) 2005 Bjarke Viksoe.
//
// Requires inclusion of the IDb.h header.
//
// The code assumes that the CSV file has a field definition header
// as the first line, which can have the following syntax:
//    name1,name2
//    int-name,"string-name"
//    ;name1,"name2"
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

#include "dblib/dblib.h"



class CCsvDatabase;
class CCsvRecordset;
class CCsvCommand;
class CCsvErrors;
class CCsvError;


class CCsvSystem : 
    public IDbSystem
  , public CComObjectRoot
  , public CComCoClass<CCsvSystem>

{
public:
   CCsvSystem();
   virtual ~CCsvSystem();

   BEGIN_COM_MAP(CCsvSystem)
   END_COM_MAP()

   BOOL Initialize();
   void Terminate();

   IDbDatabase*  CreateDatabase();
   IDbRecordset* CreateRecordset(IDbDatabase* pDb);
   IDbCommand*   CreateCommand(IDbDatabase* pDb);
};


class CCsvError : public IDbError
{
friend CCsvErrors;
friend CCsvDatabase;
protected:
   const CCsvDatabase* m_pDb;
   TCHAR m_szMsg[200];
   long m_lErrCode;

public:
   CCsvError();
   virtual ~CCsvError();

   long GetErrorCode();
   long GetNativeErrorCode();
   void GetOrigin(LPTSTR pstrStr, UINT cchMax);
   void GetSource(LPTSTR pstrStr, UINT cchMax);
   void GetMessage(LPTSTR pstrStr, UINT cchMax);
};


class CCsvErrors : public IDbErrors
{
friend CCsvDatabase;
protected:
   CCsvError m_err;
   int m_nErrors;

public:
   CCsvErrors();

   void Clear();
   long GetCount() const;
   IDbError* GetError(short iIndex);
};


class CCsvColumn
{
public:
   TCHAR szName[64];
   VARTYPE iType;
   long lOffset;
   int iSize;

public:
   CCsvColumn();
};

class CCsvDatabase : 
    public IDbDatabase
  , public CComObjectRoot
  , public CComCoClass<CCsvDatabase>

{
friend CCsvRecordset;
friend CCsvCommand;
friend CCsvError;
protected:
   TCHAR m_szFilename[MAX_PATH];
   LPSTR m_pstrText;
   char m_cSep;
   bool m_bFixedWidth;
   CCsvColumn* m_pColumns;
   short m_nCols;
   CCsvSystem* m_pSystem;
   CCsvErrors m_errs;

public:
   CCsvDatabase();
   virtual ~CCsvDatabase();

   BEGIN_COM_MAP(CCsvSystem)
   END_COM_MAP()

   void SetSystem(CCsvSystem* pSystem);


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

protected:
   BOOL _BindColumns();
   BOOL _Error(long lErrCode, LPCTSTR pstrMessage);
};

class CCsvRecordset : 
    public IDbRecordset
  , public CComObjectRoot
  , public CComCoClass<CCsvDatabase>
{
friend CCsvCommand;
protected:
   CCsvDatabase* m_pDb;
   CSimpleArray<long> m_aRows;
   CSimpleArray<long> m_aFields;
   long m_lType;
   long m_lOptions;
   long m_lCurRow;
   bool m_fAttached;

public:
   CCsvRecordset();
   virtual ~CCsvRecordset();

   void SetDatabase(CCsvDatabase* pDatabase);

   BEGIN_COM_MAP(CCsvSystem)
   END_COM_MAP()

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

// CSV specfic
public:
   BOOL MoveCursor(long lDiff, long lPos);

protected:
   BOOL _BindRows();
   BOOL _BindFields(long lRow);
   BOOL _Error(long lErrCode, LPCTSTR pstrMessage);
};

class CCsvCommand : public IDbCommand
{
protected:
   CCsvDatabase* m_pDb;
   LPTSTR m_pszSQL;
   long m_lOptions;
   DWORD m_dwRows;

public:
   CCsvCommand(CCsvDatabase* pDb);
   virtual ~CCsvCommand();

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

protected:
   BOOL _Error(long lErrCode, LPCTSTR pstrMessage);
};



#endif // !defined(AFX_CCSV_H__20011208_ED4E_9BC2_32F3_0080AD509054__INCLUDED_)
