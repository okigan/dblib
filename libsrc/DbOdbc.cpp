
#include "stdafx.h"
#include "dblib/DbOdbc.h"

#ifndef _ASSERTE
   #define _ASSERTE(x)
#endif


//////////////////////////////////////////////////////////////
// COdbcSystem
//

COdbcSystem::COdbcSystem()
   : m_henv(SQL_NULL_HENV)
{
}

COdbcSystem::~COdbcSystem()
{
   Terminate();
}

BOOL COdbcSystem::Initialize()
{
   Terminate();

   SQLRETURN rc;
   rc = ::SQLAllocHandle( SQL_HANDLE_ENV, SQL_NULL_HENV, &m_henv );
   if( rc!=SQL_SUCCESS ) return FALSE;
   rc = ::SQLSetEnvAttr( m_henv, SQL_ATTR_ODBC_VERSION, (SQLPOINTER) SQL_OV_ODBC3, SQL_IS_INTEGER );
   if( rc!=SQL_SUCCESS ) return FALSE;

   return TRUE;
}

BOOL COdbcSystem::SetConnectionPooling()
{
   SQLRETURN rc;
   rc = ::SQLSetEnvAttr( SQL_NULL_HENV, SQL_ATTR_CONNECTION_POOLING, (SQLPOINTER) SQL_CP_ONE_PER_DRIVER, 0 );
   return SQL_SUCCEEDED(rc);
}

void COdbcSystem::Terminate()
{
   if( m_henv ) {
      ::SQLFreeEnv(m_henv);
      m_henv = SQL_NULL_HENV;
   }
}

IDbDatabase* COdbcSystem::CreateDatabase()
{
  CComObject<COdbcDatabase>* obj = NULL;
  CComObject<COdbcDatabase>::CreateInstance(&obj);
  obj->SetSystem(this);

  IDbDatabase* pi = obj;

  return pi;
}

IDbRecordset* COdbcSystem::CreateRecordset(IDbDatabase* pDb)
{
  CComObject<COdbcRecordset>* obj = NULL;
  CComObject<COdbcRecordset>::CreateInstance(&obj);

  COdbcDatabase* pDatabase = dynamic_cast<COdbcDatabase*>(pDb);
  obj->SetDatabase(pDatabase);

  IDbRecordset* pi = obj;

  return pi;
}

IDbCommand* COdbcSystem::CreateCommand(IDbDatabase* pDb)
{
   return new COdbcCommand(static_cast<COdbcDatabase*>(pDb));
}


//////////////////////////////////////////////////////////////
// COdbcDatabase
//

COdbcDatabase::COdbcDatabase()
   : m_pSystem(NULL), m_hdbc(SQL_NULL_HDBC), m_hstmt(SQL_NULL_HSTMT), 
     m_lLoginTimeout(-1L), m_lQueryTimeout(-1L)
{
   ATLTRY(m_pErrors = new COdbcErrors);
#ifdef _DEBUG
   m_nRecordsets=0;
#endif
}

COdbcDatabase::~COdbcDatabase()
{
   Close();
   delete m_pErrors;
#ifdef _DEBUG
   // Check that all recordsets have been closed and deleted
   _ASSERTE(m_nRecordsets==0);
#endif
}

void COdbcDatabase::SetSystem(COdbcSystem* pSystem)
{
  m_pSystem = pSystem;
  _ASSERTE(m_pSystem);
}

BOOL COdbcDatabase::Open(HWND hWnd, LPCTSTR pstrConnectionString, LPCTSTR pstrUser, LPCTSTR pstrPassword, long iType)
{
   _ASSERTE(!::IsBadStringPtr(pstrConnectionString,(UINT)-1));
   SQLRETURN rc;

   Close();

   if( hWnd==NULL ) hWnd = ::GetActiveWindow();
   if( hWnd==NULL ) hWnd = ::GetDesktopWindow();

   // Create handle
   rc = ::SQLAllocConnect( m_pSystem->m_henv, &m_hdbc );
   if( !SQL_SUCCEEDED(rc) ) return FALSE;

   // Settings before connect...
   ::SQLSetConnectOption(m_hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_ON);
   if( m_lLoginTimeout!=-1 ) ::SQLSetConnectOption(m_hdbc, SQL_LOGIN_TIMEOUT, m_lLoginTimeout);
   if( iType & DB_OPEN_READ_ONLY ) ::SQLSetConnectOption(m_hdbc, SQL_ACCESS_MODE, SQL_MODE_READ_ONLY);

   if( iType & DB_OPEN_READ_ONLY ) {
      rc = ::SQLSetConnectOption(m_hdbc, SQL_TXN_ISOLATION, SQL_TXN_READ_UNCOMMITTED);
   }
   else {
      rc = ::SQLSetConnectOption(m_hdbc, SQL_TXN_ISOLATION, SQL_TXN_READ_COMMITTED);
   }

   if( iType & DB_OPEN_USECURSORS ) ::SQLSetConnectOption(m_hdbc, SQL_ODBC_CURSORS, SQL_CUR_USE_ODBC);
   else ::SQLSetConnectOption(m_hdbc, SQL_ODBC_CURSORS, SQL_CUR_USE_IF_NEEDED);

   //::SQLSetConnectOption(m_hdbc, SQL_OPT_TRACEFILE, (DWORD)"odbccall.txt");
   //::SQLSetConnectOption(m_hdbc, SQL_OPT_TRACE, 1);

   // Connect to database
   TCHAR szODBC[6] = { 0 };
   ::lstrcpyn(szODBC, pstrConnectionString, sizeof(szODBC)/sizeof(TCHAR));
   if( ::lstrcmpi(szODBC, _T("ODBC;") )==0 ) {
      // Connect using complex connection string
      SQLTCHAR szConnectOutput[SQL_MAX_OPTION_STRING_LENGTH+1];
      SWORD nResult;
      UWORD wConnectOption = SQL_DRIVER_COMPLETE;
      if( iType & DB_OPEN_SILENT ) wConnectOption = SQL_DRIVER_NOPROMPT;
      else if( iType & DB_OPEN_PROMPTDIALOG ) wConnectOption = SQL_DRIVER_PROMPT;
      rc = ::SQLDriverConnect(m_hdbc, hWnd, (SQLTCHAR*) pstrConnectionString, SQL_NTS, szConnectOutput, sizeof(szConnectOutput)/sizeof(SQLTCHAR), &nResult, wConnectOption);
   }
   else {
      // Connect to DSN directly
      rc = ::SQLConnect(m_hdbc, 
         (SQLTCHAR*) pstrConnectionString, (SQLSMALLINT) ::lstrlen(pstrConnectionString), 
         (SQLTCHAR*) pstrUser, (SQLSMALLINT) ::lstrlen(pstrUser), 
         (SQLTCHAR*) pstrPassword, (SQLSMALLINT) ::lstrlen(pstrPassword));
   }
   if( !SQL_SUCCEEDED(rc) ) {
      _Error();
      Close();
      return FALSE;
   }

   // Pre-allocate a statement handle for direct SQL statement...
   rc = ::SQLAllocStmt(m_hdbc, &m_hstmt);
   if( rc!=SQL_SUCCESS ) {
      _Error();
      Close();
      return FALSE;
   }

   return TRUE;
}

void COdbcDatabase::Close()
{
   if( m_hstmt ) {
      ::SQLFreeStmt(m_hstmt, SQL_DROP);
      m_hstmt = SQL_NULL_HSTMT;
   }
   if( m_hdbc ) {
      // Rollback any outstanding transactions
      ::SQLTransact(m_pSystem->m_henv, m_hdbc, SQL_ROLLBACK);
      // Disconnect database
      ::SQLDisconnect(m_hdbc);
      ::SQLFreeConnect(m_hdbc);
      m_hdbc = SQL_NULL_HDBC;
   }
}

BOOL COdbcDatabase::SetOption(SQLUSMALLINT Option, SQLULEN Value)
{
   SQLRETURN rc = ::SQLSetConnectOption(m_hdbc, Option, Value);
   return SQL_SUCCEEDED(rc) ? TRUE : _Error();
}

BOOL COdbcDatabase::SetTransactionLevel(SQLULEN Value)
{
   return SetOption(SQL_TXN_ISOLATION, Value);
}

SQLULEN COdbcDatabase::GetTransactionLevel()
{
   SQLULEN Value;
   SQLRETURN rc = ::SQLGetConnectOption(m_hdbc, SQL_TXN_ISOLATION, &Value);
   _ASSERTE(SQL_SUCCEEDED(rc)); rc;
   return Value;
}

BOOL COdbcDatabase::ExecuteSQL(LPCTSTR pstrSQL, long lType/*=DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions/*=DB_OPTION_DEFAULT*/, DWORD* pdwRowsAffected/*=NULL*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));

#ifdef _DEBUG
   ::OutputDebugString(_T("SQL: "));
   ::OutputDebugString(pstrSQL);
   ::OutputDebugString(_T("\n"));
#endif

   if( pdwRowsAffected ) *pdwRowsAffected = 0UL;

   ::SQLFreeStmt(m_hstmt, SQL_CLOSE);

   if( m_lQueryTimeout!=-1 ) ::SQLSetStmtOption(m_hstmt, SQL_QUERY_TIMEOUT, m_lQueryTimeout);
   _SetRecordsetType(m_hstmt, lType, lOptions);

   SQLRETURN rc;
   rc = ::SQLExecDirect(m_hstmt, (SQLTCHAR*) pstrSQL, SQL_NTS);   
   if( !SQL_SUCCEEDED(rc) && rc!=SQL_NO_DATA_FOUND ) return _Error(m_hstmt);

   if( pdwRowsAffected ) ::SQLRowCount(m_hstmt, (SQLINTEGER*) pdwRowsAffected);

   if( lType & DB_OPEN_TYPE_UPDATEONLY ) ::SQLFreeStmt(m_hstmt, SQL_CLOSE);

   return TRUE;
}

void COdbcDatabase::SetLoginTimeout(long lTimeout)
{
   m_lLoginTimeout = lTimeout;
}

void COdbcDatabase::SetQueryTimeout(long lTimeout)
{
   m_lQueryTimeout = lTimeout;
}

BOOL COdbcDatabase::BeginTrans()
{
   _ASSERTE(IsOpen());
   return ::SQLSetConnectOption(m_hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_OFF) == SQL_SUCCESS;
}

BOOL COdbcDatabase::CommitTrans()
{
   _ASSERTE(IsOpen());
   // Commit transaction
   SQLRETURN rc = ::SQLEndTran(SQL_HANDLE_DBC, m_hdbc, SQL_COMMIT) == SQL_SUCCESS; 
   // Turn autocommit back on again
   ::SQLSetConnectOption(m_hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_ON);
   return SQL_SUCCEEDED(rc);
}

BOOL COdbcDatabase::RollbackTrans()
{
   _ASSERTE(IsOpen());
   // Rollback transaction
   SQLRETURN rc = ::SQLEndTran(SQL_HANDLE_DBC, m_hdbc, SQL_ROLLBACK) == SQL_SUCCESS;
   // Turn autocommit back on again
   ::SQLSetConnectOption(m_hdbc, SQL_AUTOCOMMIT, SQL_AUTOCOMMIT_ON);
   return SQL_SUCCEEDED(rc);
}

BOOL COdbcDatabase::IsOpen() const
{
   return m_hdbc != SQL_NULL_HDBC;
}

IDbErrors* COdbcDatabase::GetErrors()
{
   _ASSERTE(m_pErrors);
   return m_pErrors;
}

BOOL COdbcDatabase::TranslateSQL(LPCTSTR pstrNativeSQL, LPTSTR pstrSQL, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrNativeSQL,(UINT)-1));
   _ASSERTE(!::IsBadWritePtr(pstrSQL, cchMax));
   SQLINTEGER cchSQL;
   SQLRETURN rc = ::SQLNativeSql( m_hdbc, 
      (SQLTCHAR*) pstrNativeSQL, ::lstrlen(pstrNativeSQL),
      (SQLTCHAR*) pstrSQL, (SQLINTEGER) cchMax, &cchSQL);
   return SQL_SUCCEEDED(rc);
}

BOOL COdbcDatabase::_Error(HSTMT hstmt)
{
   _ASSERTE(m_pErrors);
   m_pErrors->_Init(m_pSystem->m_henv, m_hdbc, hstmt);
   return FALSE; // Always return FALSE to allow "if( !rc ) return _Error(x)" construct.
}

void COdbcDatabase::_SetRecordsetType(HSTMT hstmt, long lType, long /*lOptions*/)
{
   // Set cursor type
   SQLUINTEGER dwCursor;
   if( lType==DB_OPEN_TYPE_DYNASET ) dwCursor = SQL_CURSOR_KEYSET_DRIVEN;
   else if( lType==DB_OPEN_TYPE_SNAPSHOT ) dwCursor = SQL_CURSOR_STATIC;
   else if( lType==DB_OPEN_TYPE_UPDATEONLY ) dwCursor = SQL_CURSOR_FORWARD_ONLY;
   else dwCursor  = SQL_CURSOR_FORWARD_ONLY;
   ::SQLSetStmtOption(hstmt, SQL_CURSOR_TYPE, dwCursor);

   // Set locking
   SQLUINTEGER dwLocking;
   if( lType==DB_OPEN_TYPE_FORWARD_ONLY ) dwLocking = SQL_CONCUR_READ_ONLY;
   else if( lType==DB_OPEN_TYPE_SNAPSHOT ) dwLocking = SQL_CONCUR_READ_ONLY;
   else if( lType==DB_OPEN_TYPE_UPDATEONLY ) dwLocking = SQL_CONCUR_READ_ONLY;
   else dwLocking = SQL_CONCUR_LOCK;
   ::SQLSetStmtOption(hstmt, SQL_CONCURRENCY, dwLocking);
}


//////////////////////////////////////////////////////////////
// COdbcRecordset
//

COdbcRecordset::COdbcRecordset()
   : m_pDb(NULL), m_hstmt(SQL_NULL_HSTMT), m_fEOF(true), 
     m_ppFields(NULL), m_nCols(0), m_fAttached(false),
     m_lQueryTimeout(-1L)
{
   _ASSERTE(m_pDb);
   _ASSERTE(m_pDb->IsOpen());
#ifdef _DEBUG
   m_pDb->m_nRecordsets++;
#endif
}

COdbcRecordset::~COdbcRecordset()
{
   Close();
#ifdef _DEBUG
   m_pDb->m_nRecordsets--;
#endif
}

void COdbcRecordset::SetDatabase(COdbcDatabase* pDb)
{
  m_pDb = pDb;
}

BOOL COdbcRecordset::Open(LPCTSTR pstrSQL, long lType, long lOptions)
{
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));

   Close();

   m_lType = lType;
   m_lOptions = lOptions;
   m_lQueryTimeout = m_pDb->m_lQueryTimeout;

   SQLRETURN rc;
   rc = ::SQLAllocStmt(m_pDb->m_hdbc, &m_hstmt);
   if( rc!=SQL_SUCCESS ) return _Error();

   COdbcDatabase::_SetRecordsetType(m_hstmt, lType, lOptions);
   if( m_lQueryTimeout!=-1 ) ::SQLSetStmtOption(m_hstmt, SQL_QUERY_TIMEOUT, m_lQueryTimeout);

#ifdef _DEBUG
   ::OutputDebugString(_T("SQL: "));
   ::OutputDebugString(pstrSQL);
   ::OutputDebugString(_T("\n"));
#endif

   // Run SQL...
   rc = ::SQLExecDirect(m_hstmt, (SQLTCHAR*) pstrSQL, SQL_NTS);
   if( !SQL_SUCCEEDED(rc) && rc!=SQL_NO_DATA_FOUND ) return _Error();

   // NOTE: Even if the query fails we still do not close it immediately.
   //       This allows the code to operate as usual and methods like IsEOF() to
   //       properly reflect the recordset state.

   if( !_BindColumns() ) return FALSE;

   return MoveNext();
}

void COdbcRecordset::Close()
{
   if( m_ppFields ) {
      for( short i=0; i<m_nCols; i++ ) delete m_ppFields[i];
      delete [] m_ppFields;
      m_ppFields = NULL;
      m_nCols = 0;
   }
   if( m_hstmt && !m_fAttached ) {
      // Give up handle
      _ASSERTE(m_pDb->IsOpen());
      ::SQLFreeStmt(m_hstmt, SQL_DROP);
   }
   m_hstmt = SQL_NULL_HSTMT;
   // Always mark as "no records" in the recordset
   m_fEOF = true;
   m_fAttached = false;
}

inline BOOL COdbcRecordset::IsOpen() const
{
   return m_hstmt != SQL_NULL_HSTMT;
}

DWORD COdbcRecordset::GetRowCount() const
{
   _ASSERTE(IsOpen());
   if( !IsOpen() ) return 0;
   SQLINTEGER lCount;
   SQLRETURN rc = ::SQLRowCount(m_hstmt, &lCount);
   return SQL_SUCCEEDED(rc) ? (DWORD) lCount : 0UL;
}

BOOL COdbcRecordset::Cancel()
{
   if( !IsOpen() ) return FALSE;
   return SQL_SUCCEEDED( ::SQLCancel(m_hstmt) );
}

BOOL COdbcRecordset::GetField(short iIndex, long& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   switch( m_ppFields[iIndex]->m_iType ) {
   case SQL_C_LONG:
   case SQL_C_ULONG:
      Data = m_ppFields[iIndex]->m_val.data.lVal;
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COdbcRecordset::GetField(short iIndex, float& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   switch( m_ppFields[iIndex]->m_iType ) {
   case SQL_C_CHAR:
      Data = (float) atof( m_ppFields[iIndex]->m_val.data.pstrVal );
      return TRUE;
   case SQL_C_FLOAT:
      Data = m_ppFields[iIndex]->m_val.data.fVal;
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COdbcRecordset::GetField(short iIndex, double& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   switch( m_ppFields[iIndex]->m_iType ) {
   case SQL_C_CHAR:
      Data = atof( m_ppFields[iIndex]->m_val.data.pstrVal );
      return TRUE;
   case SQL_C_DOUBLE:
      Data = m_ppFields[iIndex]->m_val.data.dblVal;
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COdbcRecordset::GetField(short iIndex, bool& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   switch( m_ppFields[iIndex]->m_iType ) {
   case SQL_C_LONG:
   case SQL_C_ULONG:
      Data = (bool) (m_ppFields[iIndex]->m_val.data.lVal != 0);
   case SQL_C_BIT:
      Data = m_ppFields[iIndex]->m_val.data.bVal != FALSE;
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COdbcRecordset::GetField(short iIndex, LPTSTR pData, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   switch( m_ppFields[iIndex]->m_iType ) {
   case SQL_C_CHAR:
#ifdef _UNICODE
      if( ::MultiByteToWideChar(CP_ACP, 0, m_ppFields[iIndex]->m_val.data.pstrVal, -1, pData, cchMax) == 0 ) {
         pData[0] = '\0';
         if( ::GetLastError() == ERROR_INSUFFICIENT_BUFFER ) {
            size_t cchLen = strlen(m_ppFields[iIndex]->m_val.data.pstrVal);
            LPWSTR pwstr = (LPWSTR) _alloca( (cchLen + 1) * sizeof(WCHAR) );
            ::MultiByteToWideChar(CP_ACP, 0, m_ppFields[iIndex]->m_val.data.pstrVal, -1, pwstr, cchLen + 1);
            ::lstrcpyn(pData, pwstr, cchMax);
         }
      }
#else
      ::lstrcpyn(pData, m_ppFields[iIndex]->m_val.data.pstrVal, cchMax);
#endif // _UNICODE
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)

BOOL COdbcRecordset::GetField(short iIndex, CString& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   switch( m_ppFields[iIndex]->m_iType ) {
   case SQL_C_CHAR:
      Data = m_ppFields[iIndex]->m_val.data.pstrVal;
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

#endif // __ATLSTR_H__

#if defined(_STRING_)

BOOL COdbcRecordset::GetField(short iIndex, std::string& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   switch( m_ppFields[iIndex]->m_iType ) {
   case SQL_C_CHAR:
      Data = m_ppFields[iIndex]->m_val.data.pstrVal;
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

#endif // __ATLSTR_H__

BOOL COdbcRecordset::GetField(short iIndex, SYSTEMTIME& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   ::ZeroMemory(&Data, sizeof(SYSTEMTIME));
   COdbcVariant& val = m_ppFields[iIndex]->m_val;
   switch( m_ppFields[iIndex]->m_iType ) {
   case SQL_C_TIMESTAMP:
      if( val.data.dateVal.year != 0 ) {
         Data.wYear = val.data.stampVal.year;
         Data.wMonth = val.data.stampVal.month;
         Data.wDay = val.data.stampVal.day;
         Data.wHour = val.data.stampVal.hour;
         Data.wMinute = val.data.stampVal.minute;
         Data.wSecond = val.data.stampVal.second;
      }
      return TRUE;
   case SQL_C_DATE:
      if( val.data.dateVal.year != 0 ) {
         Data.wYear = val.data.dateVal.year;
         Data.wMonth = val.data.dateVal.month;
         Data.wDay = val.data.dateVal.day;
      }
      return TRUE;
   case SQL_C_TIME:
      Data.wHour = val.data.timeVal.hour;
      Data.wMinute = val.data.timeVal.minute;
      Data.wSecond = val.data.timeVal.second;
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

DWORD COdbcRecordset::GetColumnCount() const
{
   _ASSERTE(IsOpen());
   return m_nCols;
}

DWORD COdbcRecordset::GetColumnSize(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return 0;
   return m_ppFields[iIndex]->m_dwSize;
}

short COdbcRecordset::GetColumnType(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return DB_TYPE_UNKNOWN;
   switch( m_ppFields[iIndex]->m_iSqlType ) {
   case SQL_CHAR:
   case SQL_VARCHAR:
      return DB_TYPE_CHAR;
   case SQL_INTEGER:
      return DB_TYPE_INTEGER;
   case SQL_DATE:
   case SQL_TYPE_DATE:
      return DB_TYPE_DATE;
   case SQL_TIME:
   case SQL_TYPE_TIME:
      return DB_TYPE_TIME;
   case SQL_TIMESTAMP:
   case SQL_TYPE_TIMESTAMP:
      return DB_TYPE_TIMESTAMP;
   case SQL_FLOAT:
      return DB_TYPE_REAL;
   case SQL_NUMERIC:
   case SQL_DECIMAL:
   case SQL_DOUBLE:
      return DB_TYPE_DOUBLE;
   default:
      _ASSERTE(false);
      return DB_TYPE_UNKNOWN;
   }
}

BOOL COdbcRecordset::GetColumnName(short iIndex, LPTSTR pstrName, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex<0 || iIndex>=m_nCols ) return FALSE;
   ::lstrcpyn( pstrName, m_ppFields[iIndex]->m_szName, cchMax);
   return TRUE;
}

short COdbcRecordset::GetColumnIndex(LPCTSTR pstrName) const
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrName,(UINT)-1));
   // Search for column by name.
   // NOTE: The search is case insensitive!
   for( short i=0; i<m_nCols; i++ ) {
      if( ::lstrcmp(pstrName, m_ppFields[i]->m_szName)==0 ) return i;
   }
   return -1;
}

BOOL COdbcRecordset::IsEOF() const
{
   _ASSERTE(IsOpen());
   _ASSERTE(m_nCols>0);
   return m_fEOF;
}

BOOL COdbcRecordset::MoveNext()
{
   return MoveCursor(SQL_FETCH_NEXT, 0);
}

BOOL COdbcRecordset::MovePrev()
{
   return MoveCursor(SQL_FETCH_PRIOR, 0);
}

BOOL COdbcRecordset::MoveTop()
{
   return MoveCursor(SQL_FETCH_FIRST, 0);
}

BOOL COdbcRecordset::MoveBottom()
{
   return MoveCursor(SQL_FETCH_LAST, 0);
}

BOOL COdbcRecordset::MoveAbs(DWORD dwPos)
{
   return MoveCursor(SQL_FETCH_ABSOLUTE, dwPos);
}

BOOL COdbcRecordset::MoveCursor(SHORT lType, DWORD dwPos)
{
   _ASSERTE(IsOpen());
   // Clear data values first
   for( int i=0; i<m_nCols; i++ ) m_ppFields[i]->m_val.Clear();
   // Scroll cursor
   SQLRETURN rc = lType==SQL_FETCH_NEXT ? ::SQLFetch(m_hstmt) : ::SQLFetchScroll(m_hstmt, lType, dwPos);
   // Is this the bottom of the recordset?
   m_fEOF = !(SQL_SUCCEEDED(rc));
   // Return error/success...
   return SQL_SUCCEEDED(rc) || rc==SQL_NO_DATA_FOUND ? TRUE : _Error();
}

DWORD COdbcRecordset::GetRowNumber()
{
   _ASSERTE(IsOpen());
   UDWORD rowNum;
   SQLRETURN rc = ::SQLGetStmtOption(m_hstmt, SQL_ROW_NUMBER, (SQLPOINTER) &rowNum);
   return SQL_SUCCEEDED(rc) ? (UWORD)rowNum : (DWORD)-1;
}

BOOL COdbcRecordset::NextResultset()
{
   _ASSERTE(IsOpen());
   // Get next recordset (requires multiple sql statements in query)
   SQLRETURN rc = ::SQLMoreResults(m_hstmt);
   if( !SQL_SUCCEEDED(rc) ) return _Error();
   return MoveNext();
}

BOOL COdbcRecordset::GetCursorName(LPTSTR pszName, UINT cchMax) const
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadWritePtr(pszName, cchMax));
   SQLSMALLINT cbCursor;
   SQLRETURN rc = ::SQLGetCursorName(m_hstmt, (SQLTCHAR*) pszName, (SQLSMALLINT) cchMax, &cbCursor);
   return SQL_SUCCEEDED(rc);
}

BOOL COdbcRecordset::SetCursorName(LPCTSTR pszName)
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pszName,(UINT)-1));
   SQLRETURN rc = ::SQLSetCursorName(m_hstmt, (SQLTCHAR*) pszName, (SQLSMALLINT) ::lstrlen(pszName));
   _ASSERTE(SQL_SUCCEEDED(rc));
   return SQL_SUCCEEDED(rc) ? TRUE : _Error();
}

BOOL COdbcRecordset::_BindColumns()
{
   _ASSERTE(IsOpen());

   SQLRETURN rc = ::SQLNumResultCols(m_hstmt, &m_nCols);
   _ASSERTE(SQL_SUCCEEDED(rc));
   if( m_nCols==0 ) return TRUE; // In case of INSERT/UPDATE/DELETE there
                                 // usually are no columns to bind...

   ATLTRY( m_ppFields = new COdbcField* [m_nCols] );
   if( m_ppFields==NULL ) return FALSE;

   for( SQLUSMALLINT i=0; i<m_nCols; i++ ) {
      TCHAR name[64];
      SQLSMALLINT sqltype, type, nullable, scale, namelen;       
      SQLUINTEGER len;
      rc = ::SQLDescribeCol(m_hstmt, (SQLUSMALLINT) (i+1), (SQLTCHAR*) name, sizeof(name)/sizeof(CHAR), &namelen, &sqltype, &len, &scale, &nullable);
      _ASSERTE(SQL_SUCCEEDED(rc));
      // Convert
      type = sqltype;
      _ConvertType(type, len);
      // Create Column/Field object
      ATLTRY( m_ppFields[i] = new COdbcField(name, type, sqltype, len) );
      _ASSERTE(m_ppFields[i]);
      if( m_ppFields[i]==NULL ) return FALSE;
      // Bind Field value to ODBC column
      rc = ::SQLBindCol(m_hstmt, (SQLUSMALLINT) (i+1), type, (SQLPOINTER) m_ppFields[i]->GetData(), (SQLINTEGER) len, &m_ppFields[i]->m_dwRetrivedSize);
      _ASSERTE(SQL_SUCCEEDED(rc));
   }

   return TRUE;
}

void COdbcRecordset::_ConvertType(SQLSMALLINT &type, SQLUINTEGER& len) const
{
   switch( type ) {
   case SQL_VARCHAR:
   case SQL_NUMERIC:
   case SQL_DECIMAL:
      type = SQL_C_CHAR;
      break;
   case SQL_INTEGER:
      type = SQL_C_LONG;
      len = sizeof(long);
      break;
   case SQL_FLOAT:
      type = SQL_C_FLOAT;
      len = sizeof(float);
      break;
   case SQL_DOUBLE:
      type = SQL_C_DOUBLE;
      len = sizeof(double);
      break;
   case SQL_BIT:
   case SQL_SMALLINT:
      type = SQL_C_BIT;
      len = 1;
      break;
   case SQL_TYPE_DATE:
   case SQL_DATE:
      type = SQL_C_DATE;
      len = sizeof(DATE_STRUCT);
      break;
   case SQL_TYPE_TIME:
   case SQL_TIME:
      type = SQL_C_TIME;
      len = sizeof(TIME_STRUCT);
      break;
   case SQL_TYPE_TIMESTAMP:
   case SQL_TIMESTAMP:
      type = SQL_C_TIMESTAMP;
      len = sizeof(TIMESTAMP_STRUCT);
      break;
   default:
      type = SQL_C_CHAR;
      len = 256;
   }
}

BOOL COdbcRecordset::_Error()
{
   _ASSERTE(IsOpen());
   return m_pDb->_Error(m_hstmt);
}

BOOL COdbcRecordset::_Attach(HSTMT hstmt)
{
   _ASSERTE(hstmt);
   if( m_fAttached && m_hstmt == hstmt ) {
      // Same recordset???
      // Perhaps we can optimize bind columns away here; let's try
      // it...
   }
   else {
      Close();
      m_hstmt = hstmt;
      m_fAttached = true;
      if( !_BindColumns() ) return FALSE;                                         
   }
   return MoveNext();
}


//////////////////////////////////////////////////////////////
// COdbcCommand
//

COdbcCommand::COdbcCommand(COdbcDatabase* pDb) 
   : m_pDb(pDb), m_hstmt(SQL_NULL_HSTMT), 
     m_pszSQL(NULL), m_lOptions(0), m_nParams(0), m_lQueryTimeout(-1L),
     m_bAttached(false)
{
   _ASSERTE(m_pDb);
   _ASSERTE(m_pDb->IsOpen());
#ifdef _DEBUG
   m_pDb->m_nRecordsets++;
#endif
}

COdbcCommand::~COdbcCommand()
{
   Close();
#ifdef _DEBUG
   m_pDb->m_nRecordsets--;
#endif
}

BOOL COdbcCommand::Create(LPCTSTR pstrSQL, long lType/*=DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions/*=DB_OPTION_DEFAULT*/)
{
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));

   Close();

   SQLRETURN rc;
   rc = ::SQLAllocStmt(m_pDb->m_hdbc, &m_hstmt);
   if( rc!=SQL_SUCCESS ) return _Error();

   COdbcDatabase::_SetRecordsetType(m_hstmt, lType, lOptions);

#ifdef _DEBUG
   ::OutputDebugString(_T("SQL: "));
   ::OutputDebugString(pstrSQL);
   ::OutputDebugString(_T("\n"));
#endif

   if( lOptions & DB_OPTION_PREPARE ) {
      // Prepare query plan (if driver supports it)...
      rc = ::SQLPrepare(m_hstmt, (SQLTCHAR*) pstrSQL, SQL_NTS);
      _ASSERTE(SQL_SUCCEEDED(rc));
      if( !SQL_SUCCEEDED(rc) ) return _Error();
   }
   else {
      _ASSERTE(m_pszSQL==NULL);
      m_pszSQL = new TCHAR[ ::lstrlen(pstrSQL) + 1 ];
      _ASSERTE(m_pszSQL);
      ::lstrcpy(m_pszSQL, pstrSQL);
   }

   m_nParams = 0;
   m_lOptions = lOptions;

   return TRUE;
}

BOOL COdbcCommand::Execute(IDbRecordset* pRecordset/*=NULL*/)
{
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(IsOpen());

   SQLRETURN rc = SQL_SUCCESS;

   if( m_lQueryTimeout!=-1 ) ::SQLSetStmtOption(m_hstmt, SQL_QUERY_TIMEOUT, m_lQueryTimeout);

   if( m_lOptions & DB_OPTION_PREPARE ) {
      // If command already has been execute once, we need
      // to clear the cursor state...
      if( m_bAttached ) {
         // OPTI: Presumeably closing it should be faster...
         //   while( rc!=SQL_NO_DATA ) rc = ::SQLMoreResults(m_hstmt);
         ::SQLFreeStmt(m_hstmt, SQL_CLOSE);
      }
      // Execute the prepared statememt
      rc = ::SQLExecute(m_hstmt);
   }
   else {
      // Execute unprepared SQL
      _ASSERTE(!::IsBadStringPtr(m_pszSQL,(UINT)-1));
      rc = ::SQLExecDirect(m_hstmt, (SQLTCHAR*) m_pszSQL, SQL_NTS);
   }

   // Handle data-at-execution parameters
   if( rc == SQL_NEED_DATA ) {
      SQLPOINTER pParam;
      while( (rc = ::SQLParamData(m_hstmt, &pParam))==SQL_NEED_DATA ) {
        short i=0;
         for( ; i<m_nParams; i++ ) {
            if( m_params[i].pValue==pParam ) {
               switch( m_params[i].type ) {
               case SQL_C_TIMESTAMP:
                  {
                     SYSTEMTIME* pSt = (SYSTEMTIME*) m_params[i].pValue;
                     TIMESTAMP_STRUCT& ts = m_params[i].dateVal.ts;
                     ts.year = pSt->wYear;
                     ts.month = pSt->wMonth;
                     ts.day = pSt->wDay;
                     ts.hour = pSt->wHour;
                     ts.minute = pSt->wMinute;
                     ts.second = pSt->wSecond;
                     ts.fraction = pSt->wMilliseconds;
                     rc = ::SQLPutData(m_hstmt, (SQLPOINTER) &ts, sizeof(TIMESTAMP_STRUCT));
                  }
                  break;
               case SQL_C_DATE:
                  {
                     SYSTEMTIME* pSt = (SYSTEMTIME*) m_params[i].pValue;
                     DATE_STRUCT& d = m_params[i].dateVal.date;
                     if( pSt->wYear==0 ) {
                        d.year = 1970;
                        d.month = 1;
                        d.day = 1;
                     }
                     else {
                        d.year = pSt->wYear;
                        d.month = pSt->wMonth;
                        d.day = pSt->wDay;
                     }
                     rc = ::SQLPutData(m_hstmt, (SQLPOINTER) &d, sizeof(DATE_STRUCT));
                  }
                  break;
               case SQL_C_TIME:
                  {
                     SYSTEMTIME* pSt = (SYSTEMTIME*) m_params[i].pValue;
                     TIME_STRUCT& t = m_params[i].dateVal.time;
                     t.hour = pSt->wHour;
                     t.minute = pSt->wMinute;
                     t.second = pSt->wSecond;
                     rc = ::SQLPutData(m_hstmt, (SQLPOINTER) &t, sizeof(TIME_STRUCT));
                  }
                  break;
               default:
                  _ASSERTE(!::IsBadReadPtr(m_params[i].pValue, m_params[i].size));
                  rc = ::SQLPutData(m_hstmt, (SQLPOINTER) m_params[i].pValue, m_params[i].size);
               }
               _ASSERTE(SQL_SUCCEEDED(rc));
               // Ready for next parameter...
               break;
            }
         }
         ATLASSERT(i<m_nParams);
      }     
   }

   // Check for some "safe" error-codes (e.g. SQL_NO_DATA)
   // INSERT, DELETE and UPDATE statements can cause this!
   if( rc==SQL_NO_DATA ) rc = SQL_SUCCESS;
   // If SQL failed then update error collection
   if( !SQL_SUCCEEDED(rc) ) return _Error();

   // The hstmt is kept open so we can re-execute the query
   // without creating a query-plan first (calling SQLPrepare())...

   // Did we want to see the result set?
   m_bAttached = true;
   if( pRecordset ) {
      COdbcRecordset* pRec = static_cast<COdbcRecordset*>(pRecordset);
      return pRec->_Attach(m_hstmt);
   }
   else
      return TRUE;
}

void COdbcCommand::Close()
{
   if( m_pszSQL ) {
      delete [] m_pszSQL;
      m_pszSQL = NULL;
   }
   if( m_hstmt ) {
      // Empty recordsets
      SQLRETURN rc = SQL_SUCCESS;
      while( rc != SQL_NO_DATA ) rc = ::SQLMoreResults(m_hstmt);
      // Close handle
      ::SQLFreeStmt(m_hstmt, SQL_DROP);
      m_hstmt = SQL_NULL_HSTMT;
      m_nParams = 0;
   }
}

BOOL COdbcCommand::IsOpen() const
{
   return m_hstmt != SQL_NULL_HSTMT;
}

BOOL COdbcCommand::SetOption(SQLUSMALLINT Option, SQLULEN Value)
{
   SQLRETURN rc = ::SQLSetStmtOption(m_hstmt, Option, Value);
   return SQL_SUCCEEDED(rc) ? TRUE : _Error();
}

BOOL COdbcCommand::SetCursorName(LPCTSTR pszName)
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pszName,(UINT)-1));
   SQLRETURN rc = ::SQLSetCursorName(m_hstmt, (SQLTCHAR*) pszName, (SQLSMALLINT) ::lstrlen(pszName));
   _ASSERTE(SQL_SUCCEEDED(rc));
   return SQL_SUCCEEDED(rc) ? TRUE : _Error();
}

void COdbcCommand::SetQueryTimeout(long lTimeout)
{
   m_lQueryTimeout = lTimeout;
}

DWORD COdbcCommand::GetRowCount() const
{
   _ASSERTE(IsOpen());
   if( !IsOpen() ) return 0;
   SQLINTEGER lCount;
   SQLRETURN rc = ::SQLRowCount(m_hstmt, &lCount);
   return SQL_SUCCEEDED(rc) ? (DWORD) lCount : 0UL;
}

BOOL COdbcCommand::SetParam(short iIndex, const long* pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(pData);
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   // Create entry for parameter
   m_params[iIndex].pValue = pData;
   m_params[iIndex].type = SQL_C_LONG;
   m_params[iIndex].size = sizeof(long);
   // Bind parameter
   SQLRETURN rc;
   rc = ::SQLBindParameter(m_hstmt, (SQLUSMALLINT) (iIndex+1), SQL_PARAM_INPUT, SQL_C_LONG, SQL_INTEGER, 10, 0, (SQLPOINTER) pData, 4, &m_params[iIndex].size);
   _ASSERTE(SQL_SUCCEEDED(rc));
   if( !SQL_SUCCEEDED(rc) ) return _Error();
   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short) (iIndex + 1);
   return TRUE;
}

BOOL COdbcCommand::SetParam(short iIndex, LPCTSTR pData, UINT cchMax /* = -1*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pData,(UINT)-1));
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   // Create entry for data-at-execution parameter
   m_params[iIndex].pValue = pData;
   m_params[iIndex].type = SQL_C_CHAR;
   m_params[iIndex].size = cchMax == -1 ? SQL_NTS : cchMax; // NULL terminated string;
   // Bind Parameter
   SQLRETURN rc;
#ifdef _UNICODE
   rc = ::SQLBindParameter(m_hstmt, (SQLUSMALLINT) (iIndex+1), SQL_PARAM_INPUT, SQL_C_WCHAR, SQL_CHAR, 255, 0, (SQLPOINTER) pData, 0, &m_params[iIndex].size);
#else
   rc = ::SQLBindParameter(m_hstmt, (SQLUSMALLINT) (iIndex+1), SQL_PARAM_INPUT, SQL_C_CHAR, SQL_CHAR, 255, 0, (SQLPOINTER) pData, 0, &m_params[iIndex].size);
#endif
   _ASSERTE(SQL_SUCCEEDED(rc));
   if( !SQL_SUCCEEDED(rc) ) return _Error();
   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short)(iIndex + 1);
   return TRUE;
}

#if defined(_STRING_)

BOOL COdbcCommand::SetParam(short iIndex, std::string& str)
{
   return SetParam(iIndex, str.c_str());
}

#endif // _STRING_

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)

BOOL COdbcCommand::SetParam(short iIndex, CString& str)
{
   return SetParam(iIndex, (LPCTSTR) str);
}

#endif // __ATLSTR_H__

BOOL COdbcCommand::SetParam(short iIndex, const bool* pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(pData);
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   // Create entry for parameter
   m_params[iIndex].pValue = pData;
   m_params[iIndex].type = SQL_C_BIT;
   m_params[iIndex].size = sizeof(bool);
   // Bind parameter
   // TODO: Check C-type for bool
   SQLRETURN rc;
   rc = ::SQLBindParameter(m_hstmt, (SQLUSMALLINT) (iIndex+1), SQL_PARAM_INPUT, SQL_C_BIT, SQL_BIT, 0, 0, (SQLPOINTER) pData, 0, &m_params[iIndex].size);
   _ASSERTE(SQL_SUCCEEDED(rc));
   if( !SQL_SUCCEEDED(rc) ) return _Error();
   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short)(iIndex + 1);
   return TRUE;
}

BOOL COdbcCommand::SetParam(short iIndex, const float* pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(pData);
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   // Create entry for data-at-execution parameter
   m_params[iIndex].pValue = pData;
   m_params[iIndex].type = SQL_C_FLOAT;
   m_params[iIndex].size = sizeof(float);
   // Bind parameter
   SQLRETURN rc;
   rc = ::SQLBindParameter(m_hstmt, (SQLUSMALLINT) (iIndex+1), SQL_PARAM_INPUT, SQL_C_FLOAT, SQL_FLOAT, 0, 0, (SQLPOINTER) pData, 0, &m_params[m_nParams].size);
   _ASSERTE(SQL_SUCCEEDED(rc));
   if( !SQL_SUCCEEDED(rc) ) return _Error();
   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short)(iIndex + 1);
   return TRUE;
}

BOOL COdbcCommand::SetParam(short iIndex, const double* pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(pData);
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   // Create entry for parameter
   m_params[iIndex].pValue = pData;
   m_params[iIndex].type = SQL_C_DOUBLE;
   m_params[iIndex].size = sizeof(double);
   // Bind parameter
   SQLRETURN rc;
   rc = ::SQLBindParameter(m_hstmt, (SQLUSMALLINT) (iIndex+1), SQL_PARAM_INPUT, SQL_C_DOUBLE, SQL_DOUBLE, 0, 0, (SQLPOINTER) pData, 0, &m_params[iIndex].size);
   _ASSERTE(SQL_SUCCEEDED(rc));
   if( !SQL_SUCCEEDED(rc) ) return _Error();
   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short)(iIndex + 1);
   return TRUE;
}

BOOL COdbcCommand::SetParam(short iIndex, const SYSTEMTIME* pData, short iType)
{
   _ASSERTE(IsOpen());
   _ASSERTE(pData);
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   // Figure out if it's a date/time/timestamp
   SQLSMALLINT type;
   short ctype;
   DWORD len;
   DWORD size = (DWORD) SQL_DATA_AT_EXEC;
   switch( iType ) {
   case DB_TYPE_TIME:
      type = SQL_TIME;
      ctype = SQL_C_TIME;
      len = sizeof(TIME_STRUCT);
      break;
   case DB_TYPE_DATE:
      type = SQL_DATE;
      ctype = SQL_C_DATE;
      len = sizeof(DATE_STRUCT);
      if( pData->wYear == 0 ) size = (DWORD) SQL_NULL_DATA;
      break;
   default:
      type = SQL_TIMESTAMP;
      ctype = SQL_C_TIMESTAMP;
      len = sizeof(TIMESTAMP_STRUCT);
   }
   // Create entry for "data-at-execution" parameter
   // Must be at execution time because we're converting the date structure
   // ourselves.
   m_params[iIndex].pValue = pData;
   m_params[iIndex].type = ctype;
   m_params[iIndex].size = size;
   // Bind it...
   SQLRETURN rc;
   rc = ::SQLBindParameter(m_hstmt, (SQLUSMALLINT) (iIndex+1), SQL_PARAM_INPUT, ctype, type, 0, 0, (SQLPOINTER) pData, 0, &m_params[iIndex].size);
   _ASSERTE(SQL_SUCCEEDED(rc));
   if( !SQL_SUCCEEDED(rc) ) return _Error();
   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short)(iIndex + 1);
   return TRUE;
}

BOOL COdbcCommand::_Error()
{
   _ASSERTE(m_pDb);
   return m_pDb->_Error(m_hstmt);
}


//////////////////////////////////////////////////////////////
// COdbcField
//

COdbcField::COdbcField(LPCTSTR pstrColName, SQLSMALLINT iType, SQLSMALLINT iSqlType, SQLINTEGER dwSize)
{
   _ASSERTE(pstrColName);
   _ASSERTE(::lstrlen(pstrColName)<sizeof(m_szName)/sizeof(TCHAR)); 
   ::lstrcpyn(m_szName, pstrColName, sizeof(m_szName)/sizeof(TCHAR));
   m_iType = iType;
   m_iSqlType = iSqlType;
   m_dwSize = dwSize;
   m_val.Init(iType, dwSize);
}

COdbcField::~COdbcField()
{
}

LPVOID COdbcField::GetData() const
{
   switch( m_iType ) {
   case SQL_C_CHAR:
      return (LPVOID) m_val.data.pstrVal;
   default:
      return (LPVOID) &m_val.data;
   }
}


//////////////////////////////////////////////////////////////
// COdbcVariant
//

COdbcVariant::COdbcVariant()
   : m_iType(SQL_TYPE_NULL)
{
}

COdbcVariant::~COdbcVariant()
{
   Destroy();
}

void COdbcVariant::Init(SHORT iType, DWORD dwSize)
{
   m_iType = iType;
   m_dwSize = dwSize;
   switch( m_iType ) {
   case SQL_C_CHAR:
      m_dwSize++;
      data.pstrVal = new CHAR[m_dwSize];
      break;
   }
}

void COdbcVariant::Destroy()
{
   switch( m_iType ) {
   case SQL_C_CHAR:
      delete [] data.pstrVal;
      break;
   }
}

void COdbcVariant::Clear()
{
   switch( m_iType ) {
   case SQL_C_CHAR:
      ::ZeroMemory(data.pstrVal, m_dwSize);
      break;
   default:
      data.lVal = 0;
   }
}


//////////////////////////////////////////////////////////////
// COdbcErrors
//

COdbcErrors::COdbcErrors() 
   : m_lCount(0L)
{
}

COdbcErrors::~COdbcErrors()
{
}

void COdbcErrors::_Init(HENV henv, HDBC hdbc, HSTMT hstmt)
{
   TCHAR szMsg[SQL_MAX_MESSAGE_LENGTH+1];
   TCHAR szState[SQL_SQLSTATE_SIZE+1];
   long lNative = 0;
   SQLSMALLINT iCount = 0;
   short i = 0;
   while( SQL_SUCCEEDED( ::SQLError(henv, hdbc, hstmt, (SQLTCHAR*) szState, &lNative, (SQLTCHAR*) szMsg, sizeof(szMsg)/sizeof(TCHAR), &iCount) ) ) {
#if defined(_DEBUG) && !defined(_NO_ODBC_WARNINGS)
      ::MessageBox(NULL, szMsg, _T("ODBC Error"), MB_OK|MB_ICONERROR|MB_SYSTEMMODAL);
#endif
      m_p[i]._Init( (LPTSTR) szState, lNative, (LPTSTR) szMsg );
      i++;
      if( i>=MAX_ERRORS ) break;
   }
   m_lCount = i;
}

long COdbcErrors::GetCount() const
{
   return m_lCount;
}

void COdbcErrors::Clear()
{
   m_lCount = 0;
}

IDbError* COdbcErrors::GetError(short iIndex)
{
   _ASSERTE(iIndex>=0 && iIndex<m_lCount);
   if( iIndex<0 || iIndex>=m_lCount ) return NULL;
   return &m_p[iIndex];
}


//////////////////////////////////////////////////////////////
// COdbcError
//

COdbcError::COdbcError()
{
}

COdbcError::~COdbcError()
{
}

void COdbcError::_Init(LPCTSTR pstrState, long lNative, LPCTSTR pstrMsg)
{
   _ASSERTE(::lstrlen(pstrState)<sizeof(m_szState)/sizeof(TCHAR));
   _ASSERTE(::lstrlen(pstrMsg)<sizeof(m_szMsg)/sizeof(TCHAR));
   m_lNative = lNative;
   ::lstrcpyn(m_szState, pstrState, sizeof(m_szState)/sizeof(TCHAR));
   ::lstrcpyn(m_szMsg, pstrMsg, sizeof(m_szMsg)/sizeof(TCHAR));
}

long COdbcError::GetErrorCode()
{
   return 0;
}

long COdbcError::GetNativeErrorCode()
{
   return m_lNative;
}

void COdbcError::GetState(LPTSTR pstrStr, UINT cchMax)
{
   ::lstrcpyn(pstrStr, m_szState, cchMax);
}

void COdbcError::GetOrigin(LPTSTR pstrStr, UINT cchMax)
{
   // Find end of ODBC headers
   LPCTSTR p = m_szMsg;
   while( *p==_T('[') ) {
      while( *p && *p != _T(']') ) p = ::CharNext(p);
      if( *p ) p = ::CharNext(p);
   }
   UINT len = (DWORD) p - (DWORD) m_szMsg;
   ::lstrcpyn(pstrStr, m_szMsg, min(len, cchMax));
}

void COdbcError::GetMessage(LPTSTR pstrStr, UINT cchMax)
{
   // Skip ODBC headers and leading spaces (MS-Access thing)...
   LPCTSTR p = m_szMsg;
   while( *p==_T('[') ) {
      while( *p && *p != _T(']') ) p = ::CharNext(p);
      if( *p ) p = ::CharNext(p);
      while( *p && *p == _T(' ') ) p = ::CharNext(p);
   }
   // Copy actual message
   ::lstrcpyn(pstrStr, p, cchMax);
}

void COdbcError::GetSource(LPTSTR pstrStr, UINT cchMax)
{
   ::lstrcpyn(pstrStr, m_szMsg, cchMax);
}
