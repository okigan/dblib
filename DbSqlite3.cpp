
#include "stdafx.h"
#include "DbSqlite3.h"

#ifndef _ASSERTE
   #define _ASSERTE(x)
#endif


//////////////////////////////////////////////////////////////
// CSqlite3System
//

BOOL CSqlite3System::Initialize()
{   
   return TRUE; // No action required!
}

void CSqlite3System::Terminate()
{
}

IDbDatabase*  CSqlite3System::CreateDatabase()
{
   return new CSqlite3Database(this);
}

IDbRecordset* CSqlite3System::CreateRecordset(IDbDatabase* pDb)
{
   return new CSqlite3Recordset(reinterpret_cast<CSqlite3Database*>(pDb));
}

IDbCommand* CSqlite3System::CreateCommand(IDbDatabase* pDb)
{
   return new CSqlite3Command(reinterpret_cast<CSqlite3Database*>(pDb));
}


//////////////////////////////////////////////////////////////
// CSqlite3Error
//

long CSqlite3Error::GetErrorCode()
{
   return m_iError;
}

long CSqlite3Error::GetNativeErrorCode()
{
   return m_iNative;
}

void CSqlite3Error::GetOrigin(LPTSTR pstrStr, UINT cchMax)
{
   _tcsncpy(pstrStr, _T("SQLite3"), cchMax);
}

void CSqlite3Error::GetSource(LPTSTR pstrStr, UINT cchMax)
{
   _tcsncpy(pstrStr, _T("SQLite3"), cchMax);
}

void CSqlite3Error::GetMessage(LPTSTR pstrStr, UINT cchMax)
{
   USES_CONVERSION;
   _tcsncpy(pstrStr, A2CT(m_szMsg), cchMax);
}

void CSqlite3Error::_Init(int iError, int iNative, LPCSTR pstrMsg)
{
   m_iError = iError;
   m_iNative = iNative;
   strncpy(m_szMsg, pstrMsg, sizeof(m_szMsg));
   m_szMsg[ sizeof(m_szMsg) - 1 ] = '\0';
#if defined(WIN32) && defined(_DEBUG)
   ::MessageBoxA(NULL, m_szMsg, "SQL Error", MB_ICONERROR | MB_SETFOREGROUND);
#endif
}


//////////////////////////////////////////////////////////////
// CSqlite3Errors
//

void CSqlite3Errors::Clear()
{
   m_lCount = 0;
}

long CSqlite3Errors::GetCount() const
{
   return m_lCount;
}

IDbError* CSqlite3Errors::GetError(short iIndex)
{
   if( iIndex != 0 ) return NULL;
   return &m_Error;
}

void CSqlite3Errors::_Init(int iError, int iNative, LPCSTR pstrMessage)
{
   m_Error._Init(iError, iNative, pstrMessage);
}


//////////////////////////////////////////////////////////////
// CSqlite3Database
//

CSqlite3Database::CSqlite3Database(CSqlite3System* pSystem) : m_pSystem(pSystem), m_pDb(NULL)
{
}

CSqlite3Database::~CSqlite3Database()
{
   Close();
}

BOOL CSqlite3Database::Open(HWND /*hWnd*/, LPCTSTR pstrConnectionString, LPCTSTR /*pstrUser*/, LPCTSTR /*pstrPassword*/, long /*iType = DB_OPEN_DEFAULT*/)
{
   _ASSERTE(m_pSystem);
   _ASSERTE(pstrConnectionString);
   // Sqlite does not fail if the database doesn't exists; it creates it!
   if( ::GetFileAttributes(pstrConnectionString) == (DWORD) -1 ) return _Error(99, SQLITE_NOTFOUND, "Database does not exist");
   Close();
   // Open it...
   USES_CONVERSION;
   int iErr = ::sqlite3_open(T2CA(pstrConnectionString), &m_pDb);
   if( iErr != SQLITE_OK ) {
      if( m_pDb != NULL ) return _Error(::sqlite3_errcode(m_pDb), iErr, ::sqlite3_errmsg(m_pDb));
      return _Error(1, iErr, "Unable to open database");
   }
   ::sqlite3_busy_timeout(m_pDb, 10000L);
   return TRUE;
}

void CSqlite3Database::Close()
{
   if( m_pDb == NULL ) return;
   ::sqlite3_close(m_pDb);
   m_pDb = NULL;
}

BOOL CSqlite3Database::IsOpen() const
{
   return m_pDb != NULL;
}

BOOL CSqlite3Database::ExecuteSQL(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/, DWORD* pdwRowsAffected /*= NULL*/)
{
   CSqlite3Command cmd(this);
   if( !cmd.Create(pstrSQL, lType, lOptions) ) return FALSE;
   if( !cmd.Execute() ) return FALSE;
   if( pdwRowsAffected ) *pdwRowsAffected = cmd.GetRowCount();
   return TRUE;
}

BOOL CSqlite3Database::BeginTrans()
{
   return ExecuteSQL(_T("BEGIN TRANSACTION"));
}

BOOL CSqlite3Database::CommitTrans()
{
   return ExecuteSQL(_T("COMMIT TRANSACTION"));
}

BOOL CSqlite3Database::RollbackTrans()
{
   return ExecuteSQL(_T("ROLLBACK TRANSACTION"));
}

void CSqlite3Database::SetLoginTimeout(long /*lTimeout*/)
{
}

void CSqlite3Database::SetQueryTimeout(long lTimeout)
{
   ::sqlite3_busy_timeout(m_pDb, lTimeout);
}

IDbErrors* CSqlite3Database::GetErrors()
{
   return &m_Errors;
}

INT64 CSqlite3Database::GetLastRowId()
{
   return ::sqlite3_last_insert_rowid(m_pDb);
}

BOOL CSqlite3Database::_Error(int iError, int iNative, LPCSTR pstrMessage)
{   
   m_Errors._Init(iError, iNative, pstrMessage);
   return FALSE; // Always returns FALSE!
}


//////////////////////////////////////////////////////////////
// CSqlite3Recordset
//

CSqlite3Recordset::CSqlite3Recordset(CSqlite3Database* pDb) : m_pDb(pDb), m_pVm(NULL), m_ppSnapshot(NULL), m_bAttached(false)
{
}

CSqlite3Recordset::~CSqlite3Recordset()
{
   Close();
}

BOOL CSqlite3Recordset::Open(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/)
{
   _ASSERTE(m_pDb);
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));

   USES_CONVERSION;
#if _ATL_VER > 0x0700 && WINVER >= 0x0500
   _acp = CP_UTF8;
#endif

   Close();

   m_ppSnapshot = NULL;
   m_pVm = NULL;
   m_nCols = 0;
   m_nRows = 0;
   m_fEOF = TRUE;
   m_lType = lType;
   m_lOptions = lOptions;

   switch( lType ) {
   case DB_OPEN_TYPE_FORWARD_ONLY:
      {
         LPCSTR pTrail = NULL;
         LPCSTR pstrZSQL = T2CA(pstrSQL);
         int iErr = ::sqlite3_prepare(*m_pDb, pstrZSQL, strlen(pstrZSQL), &m_pVm, &pTrail);
         if( iErr != SQLITE_OK ) return _Error(::sqlite3_errcode(*m_pDb), iErr, ::sqlite3_errmsg(*m_pDb));
         _ASSERTE(strlen(pTrail)==0);  // We don't really support batch SQL statements
         m_iPos = -1;
         m_fEOF = FALSE;
         m_nCols = ::sqlite3_column_count(m_pVm);
         MoveNext();
         return TRUE;
      }
      break;
   default:
      {
         LPSTR pstrErrorMsg = NULL;
         int iErr = ::sqlite3_get_table(*m_pDb, T2CA(pstrSQL), &m_ppSnapshot, &m_nRows, &m_nCols, &pstrErrorMsg);
         if( iErr != SQLITE_OK ) {
            _Error(::sqlite3_errcode(*m_pDb), iErr, pstrErrorMsg);
            if( pstrErrorMsg != NULL ) ::sqlite3_free(pstrErrorMsg);
            return FALSE;
         }
         m_iPos = 0;
         return TRUE;
      }
   }
}

void CSqlite3Recordset::Close()
{
   if( m_ppSnapshot ) {
      ::sqlite3_free_table(m_ppSnapshot);
      m_ppSnapshot = NULL;
   }
   if( m_pVm && !m_bAttached ) {
      ::sqlite3_finalize(m_pVm);
      m_pVm = NULL;
   }
   m_bAttached = false;
}

BOOL CSqlite3Recordset::IsOpen() const
{
   return (m_ppSnapshot != NULL) || (m_pVm != NULL);
}

DWORD CSqlite3Recordset::GetRowCount() const
{
   _ASSERTE(IsOpen());
   return ::sqlite3_changes(*m_pDb);
}

BOOL CSqlite3Recordset::GetField(short iIndex, long& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      Data = ::sqlite3_column_int(m_pVm, iIndex);
   }
   else {
      LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
      if( pstr == NULL ) {
         Data = 0;
         return TRUE;
      }
      else {
         Data = atol(pstr);
      }
   }
   return TRUE;
}

BOOL CSqlite3Recordset::GetField(short iIndex, float& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      Data = (float) ::sqlite3_column_double(m_pVm, iIndex);
   }
   else {
      LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
      if( pstr == NULL ) {
         Data = 0.0f;
         return TRUE;
      }
      else {
         Data = (float) atof(pstr);
      }
   }
   return TRUE;
}

BOOL CSqlite3Recordset::GetField(short iIndex, double& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      Data = ::sqlite3_column_double(m_pVm, iIndex);
   }
   else {
      LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
      if( pstr == NULL ) {
         Data = 0.0;
         return TRUE;
      }
      else {
         Data = atof(pstr);
      }
   }
   return TRUE;
}

BOOL CSqlite3Recordset::GetField(short iIndex, bool& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPSTR pstr = NULL;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      pstr = (LPSTR) ::sqlite3_column_text(m_pVm, iIndex);
   }
   else {
      pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   }
   if( pstr == NULL ) return FALSE;
   // 1, -1, True, true, Yes, yes
   Data = *pstr == '1' || *pstr == '-' || *pstr == 'T' || *pstr == 't' || *pstr == 'Y' || *pstr == 'y';
   return TRUE;
}

BOOL CSqlite3Recordset::GetField(short iIndex, LPTSTR pData, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
#if !defined(UNICODE)
      USES_CONVERSION;
      _tcsncpy(pData, A2T( (char*)::sqlite3_column_text(m_pVm, iIndex) ), cchMax);
#else  // UNICODE
      _tcsncpy(pData, (WCHAR*) ::sqlite3_column_text16(m_pVm, iIndex), cchMax);
#endif // UNICODE
   }
   else {
      LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
      if( pstr == NULL ) {
         _tcscpy(pData, _T(""));
      }
      else {
         USES_CONVERSION;
         LPCSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
         _tcsncpy(pData, A2CT(pstr), cchMax);
      }
   }
   return TRUE;
}

BOOL CSqlite3Recordset::GetField(short iIndex, SYSTEMTIME& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   ::ZeroMemory(&Data, sizeof(SYSTEMTIME));
   LPSTR pstr = NULL;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      pstr = (LPSTR) ::sqlite3_column_text(m_pVm, iIndex);
   }
   else {
      pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   }
   if( pstr == NULL ) return FALSE;
   char szDate[40] = { 0 };
   pstr = strncpy(szDate, pstr, sizeof(szDate) - 1);
   int nLen = strlen(szDate);
   if( nLen == 8 ) {
      pstr[2] = '\0';
      pstr[5] = '\0';
      Data.wHour = (WORD) atoi(&pstr[0]);
      Data.wMinute = (WORD) atoi(&pstr[3]);
      Data.wSecond = (WORD) atoi(&pstr[6]);
   }
   if( nLen >= 10 ) {
      pstr[4] = '\0';
      pstr[7] = '\0';
      pstr[10] = '\0';
      Data.wYear = (WORD) atoi(&pstr[0]);
      Data.wMonth = (WORD) atoi(&pstr[5]);
      Data.wDay = (WORD) atoi(&pstr[8]);
   }
   if( nLen >= 19 ) {
      pstr[13] = '\0';
      pstr[16] = '\0';
      pstr[19] = '\0';
      Data.wHour = (WORD) atoi(&pstr[11]);
      Data.wMinute = (WORD) atoi(&pstr[14]);
      Data.wSecond = (WORD) atoi(&pstr[17]);
   }
   if( nLen > 20 ) {
      Data.wMilliseconds = (WORD) atoi(&pstr[20]);
   }
   return TRUE;
}

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)

BOOL CSqlite3Recordset::GetField(short iIndex, CString& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
#if !defined(UNICODE)
      Data = (char*) ::sqlite3_column_text(m_pVm, iIndex);
#else  // UNICODE
      Data = ::sqlite3_column_text16(m_pVm, iIndex);
#endif // UNICODE
   }
   else {
      LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
      if( pstr == NULL ) {
         Data = _T("");
      }
      else {
         Data = pstr;
      }
   }
   return TRUE;
}

#endif // __ATLSTR_H__

#if defined(_STRING_)

BOOL CSqlite3Recordset::GetField(short iIndex, std::string& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      Data = (char*) ::sqlite3_column_text(m_pVm, iIndex);
   }
   else {
      LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
      if( pstr == NULL ) {
         Data = "";
      }
      else {
         Data = pstr;
      }
   }
   return TRUE;
}

#endif // __ATLSTR_H__

DWORD CSqlite3Recordset::GetColumnSize(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex < 0 || iIndex >= m_nCols ) return 0;
   return 255;  // Hardcoded value! Sqlite actually supports up to many Mb of data!
}

BOOL CSqlite3Recordset::GetColumnName(short iIndex, LPTSTR pstrName, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   USES_CONVERSION;
   LPCSTR pstrSrc = m_lType == DB_OPEN_TYPE_FORWARD_ONLY ? ::sqlite3_column_name(m_pVm, iIndex) : m_ppSnapshot[iIndex];
   _tcsncpy(pstrName, A2CT(pstrSrc), cchMax);
   return TRUE;
}

short CSqlite3Recordset::GetColumnType(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPCSTR pstr = NULL;
   if( m_lType != DB_OPEN_TYPE_FORWARD_ONLY ) return DB_TYPE_UNKNOWN;
   pstr = ::sqlite3_column_decltype(m_pVm, iIndex);
   if( pstr == NULL ) return DB_TYPE_UNKNOWN;
      if( strnicmp(pstr, "INT", 3) == 0 ) return DB_TYPE_INTEGER;
      if( strnicmp(pstr, "DATE", 4) == 0 ) return DB_TYPE_DATE;
      if( strnicmp(pstr, "TIME", 4) == 0 ) return DB_TYPE_TIME;
      if( strnicmp(pstr, "BOOL", 4) == 0 ) return DB_TYPE_BOOLEAN;
      if( strnicmp(pstr, "TEXT", 4) == 0 ) return DB_TYPE_CHAR;
      if( strnicmp(pstr, "CHAR", 4) == 0 ) return DB_TYPE_CHAR;
      if( strnicmp(pstr, "FLOAT", 5) == 0 ) return DB_TYPE_DOUBLE;
      if( strnicmp(pstr, "DOUBLE", 6) == 0 ) return DB_TYPE_DOUBLE;
      if( strnicmp(pstr, "VARCHAR", 7) == 0 ) return DB_TYPE_CHAR;
      if( strnicmp(pstr, "NUMERIC", 7) == 0 ) return DB_TYPE_DOUBLE;
      if( strnicmp(pstr, "SMALLINT", 8) == 0 ) return DB_TYPE_INTEGER;
      if( strnicmp(pstr, "TIMESTAMP", 9) == 0 ) return DB_TYPE_TIMESTAMP;
   return DB_TYPE_UNKNOWN;
}

short CSqlite3Recordset::GetColumnIndex(LPCTSTR pstrName) const
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrName,(UINT)-1));
   USES_CONVERSION;
   LPCSTR pstr = T2CA(pstrName);
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      for( short iIndex = 0; iIndex < m_nCols; iIndex++ ) if( strcmp(pstr, ::sqlite3_column_name(m_pVm, iIndex)) == 0 ) return iIndex;
   }
   else {
      for( short iIndex = 0; iIndex < m_nCols; iIndex++ ) if( strcmp(pstr, m_ppSnapshot[iIndex]) == 0 ) return iIndex;
   }
   return -1;
}

DWORD CSqlite3Recordset::GetColumnCount() const
{
   _ASSERTE(IsOpen());
   return m_nCols;
}

BOOL CSqlite3Recordset::IsEOF() const
{
   _ASSERTE(IsOpen());
   if( !IsOpen() ) return TRUE;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return m_fEOF;
   return m_iPos >= m_nRows;
}

BOOL CSqlite3Recordset::MoveNext()
{   
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      int iErr = ::sqlite3_step(m_pVm);
      m_fEOF = iErr != SQLITE_ROW;
      _ASSERTE(iErr==SQLITE_DONE || iErr==SQLITE_ROW);
      if( iErr != SQLITE_DONE && iErr != SQLITE_ROW ) return _Error(1, iErr, "Move Error");
      return TRUE;
   }
   else {
      _ASSERTE(IsOpen());
      m_iPos++;
      return IsEOF();
   }
}

BOOL CSqlite3Recordset::MovePrev()
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return _Error(1, 1, "Invalid recordset type");
   if( m_iPos <= 0 ) return FALSE;
   --m_iPos;
   return TRUE;
}

BOOL CSqlite3Recordset::MoveTop()
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return _Error(1, 1, "Invalid recordset type");
   m_iPos = 0;
   return TRUE;
}

BOOL CSqlite3Recordset::MoveBottom()
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return _Error(1, 1, "Invalid recordset type");
   if( m_nRows == 0 ) return FALSE;
   m_iPos = m_nRows - 1;
   return IsEOF();
}

BOOL CSqlite3Recordset::MoveAbs(DWORD dwPos)
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return _Error(1, 1, "Invalid recordset type");
   if( m_iPos < 0 ) return FALSE;
   if( m_iPos >= m_nRows ) return FALSE;
   m_iPos = (int) dwPos;
   return TRUE;
}

DWORD CSqlite3Recordset::GetRowNumber()
{
   _ASSERTE(IsOpen());
   _ASSERTE(m_lType!=DB_OPEN_TYPE_FORWARD_ONLY);
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return 0;
   return m_iPos;
}

BOOL CSqlite3Recordset::NextResultset()
{
   _ASSERTE(IsOpen());
   _ASSERTE(false);
   return FALSE;
}

BOOL CSqlite3Recordset::_Error(int iError, int iNative, LPCSTR pstrMessage)
{
   m_pDb->_Error(iError, iNative, pstrMessage);
   return FALSE;
}

BOOL CSqlite3Recordset::_Attach(sqlite3_stmt* pVm)
{
   Close();
   m_pVm = pVm;
   m_ppSnapshot = NULL;
   m_lType = DB_OPEN_TYPE_FORWARD_ONLY;
   m_lOptions = 0;
   m_nCols = ::sqlite3_column_count(pVm);
   m_nRows = 0;
   m_iPos = -1;
   m_fEOF = FALSE;
   m_bAttached = true;
   return MoveNext();
}


//////////////////////////////////////////////////////////////
// CSqlite3Command
//

CSqlite3Command::CSqlite3Command(CSqlite3Database* pDb) : m_pDb(pDb), m_pVm(NULL)
{
}

CSqlite3Command::~CSqlite3Command()
{
   Close();
}

BOOL CSqlite3Command::Create(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long /*lOptions = DB_OPTION_DEFAULT*/)
{
   _ASSERTE(m_pDb);
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));
   _ASSERTE(lType==DB_OPEN_TYPE_FORWARD_ONLY);

   Close();

   m_nParams = 0;
   if( lType != DB_OPEN_TYPE_FORWARD_ONLY ) return FALSE;

   USES_CONVERSION;
   LPCSTR pTrail = NULL;
   LPCSTR pstrZSQL = T2CA(pstrSQL);
   int iErr = ::sqlite3_prepare(*m_pDb, pstrZSQL, strlen(pstrZSQL), &m_pVm, &pTrail);
   if( iErr != SQLITE_OK ) return _Error(::sqlite3_errcode(*m_pDb), iErr, ::sqlite3_errmsg(*m_pDb));

   return TRUE;
}

BOOL CSqlite3Command::Execute(IDbRecordset* pRecordset /*= NULL*/)
{
   int iErr = ::sqlite3_reset(m_pVm);
   if( iErr != SQLITE_OK ) return _Error(::sqlite3_errcode(*m_pDb), iErr, ::sqlite3_errmsg(*m_pDb));
   
   LPCSTR pstr = NULL;
   for( int iIndex = 0; iIndex < m_nParams; iIndex++ ) {
      TSqlite3Param& Param = m_Params[iIndex];
      switch( Param.vt ) {
      case VT_I4:
         iErr = ::sqlite3_bind_int(m_pVm, iIndex + 1, * (long*) Param.pVoid);
         break;
      case VT_R4:
         iErr = ::sqlite3_bind_double(m_pVm, iIndex + 1, * (float*) Param.pVoid);
         break;
      case VT_R8:
         iErr = ::sqlite3_bind_double(m_pVm, iIndex + 1, * (double*) Param.pVoid);
         break;
      case VT_LPSTR:
         pstr = (const char*) Param.pVoid;
         if( Param.len == -2 ) iErr = ::sqlite3_bind_null(m_pVm, iIndex + 1);
         else if( Param.len == -1 ) iErr = ::sqlite3_bind_text(m_pVm, iIndex + 1, pstr, strlen(pstr), SQLITE_STATIC);
         else iErr = ::sqlite3_bind_text(m_pVm, iIndex + 1, pstr, Param.len, SQLITE_STATIC);
         break;
      case VT_BOOL:
         iErr = ::sqlite3_bind_int(m_pVm, iIndex + 1, (* (bool*) Param.pVoid) & 1);
         break;
#if defined(_STRING_)
      case VT_BSTR:
         pstr = (const char*) ((std::string*) Param.pVoid)->c_str();
         iErr = ::sqlite3_bind_text(m_pVm, iIndex + 1, pstr, ((std::string*) Param.pVoid)->length(), SQLITE_STATIC);
         break;
#endif
#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)
      case VT_BSTR_BLOB:
         iErr = ::sqlite3_bind_text(m_pVm, iIndex + 1, (LPCTSTR) * ((CString*) Param.pVoid), ((CString*) Param.pVoid)->GetLength(), SQLITE_STATIC);
         break;
#endif
      case VT_DATE:
         {
            SYSTEMTIME* pData = (SYSTEMTIME*) Param.pVoid;
            CHAR szBuffer[32];
            int iLen = ::wsprintfA(szBuffer, "%04ld-%02ld-%02ld %02ld:%02ld:%02ld.%ld", 
               (long) pData->wYear,
               (long) pData->wMonth,
               (long) pData->wDay,
               (long) pData->wHour,
               (long) pData->wMinute,
               (long) pData->wSecond,
               (long) pData->wMilliseconds);
            if( Param.len == DB_TYPE_DATE ) iLen = 10; // date part
            if( pData->wYear == 0 ) iErr = ::sqlite3_bind_null(m_pVm, iIndex + 1);
            else iErr = ::sqlite3_bind_text(m_pVm, iIndex + 1, szBuffer, iLen, NULL);
         }
         break;
      }
      if( iErr != SQLITE_OK ) return _Error(::sqlite3_errcode(*m_pDb), iErr, ::sqlite3_errmsg(*m_pDb));
   }

   if( pRecordset ) {
      CSqlite3Recordset* pRec = reinterpret_cast<CSqlite3Recordset*>(pRecordset);
      return pRec->_Attach(m_pVm);
   }
   else {
      CSqlite3Recordset rec = m_pDb;
      return rec._Attach(m_pVm);
   }
}

void CSqlite3Command::Close()
{
   if( m_pVm ) {
      ::sqlite3_finalize(m_pVm);
      m_pVm = NULL;
   }
}

BOOL CSqlite3Command::IsOpen() const
{
   return m_pVm != NULL;
}

DWORD CSqlite3Command::GetRowCount() const
{
   _ASSERTE(IsOpen());
   return ::sqlite3_changes(*m_pDb);
}

BOOL CSqlite3Command::SetParam(short iIndex, const long* pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_I4;
   m_Params[iIndex].pVoid = pData;
   m_Params[iIndex].len = sizeof(long);
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

BOOL CSqlite3Command::SetParam(short iIndex, LPCTSTR pData, UINT cchMax /*= -1*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pData,(UINT)-1));
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_LPSTR;
   m_Params[iIndex].pVoid = pData;
   m_Params[iIndex].len = (int) cchMax;
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

BOOL CSqlite3Command::SetParam(short iIndex, const bool* pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_BOOL;
   m_Params[iIndex].pVoid = pData;
   m_Params[iIndex].len = sizeof(bool);
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

BOOL CSqlite3Command::SetParam(short iIndex, const float* pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_R4;
   m_Params[iIndex].pVoid = pData;
   m_Params[iIndex].len = sizeof(float);
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

BOOL CSqlite3Command::SetParam(short iIndex, const double* pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_R8;
   m_Params[iIndex].pVoid = pData;
   m_Params[iIndex].len = sizeof(double);
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

BOOL CSqlite3Command::SetParam(short iIndex, const SYSTEMTIME* pData, short iType /*= DB_TYPE_TIMESTAMP*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_DATE;
   m_Params[iIndex].pVoid = pData;
   m_Params[iIndex].len = iType;
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

#if defined(_STRING_)

BOOL CSqlite3Command::SetParam(short iIndex, std::string& str)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_BSTR;
   m_Params[iIndex].pVoid = &str;
   m_Params[iIndex].len = -1;
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

#endif // _STRING

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)

BOOL CSqlite3Command::SetParam(short iIndex, CString& str)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_BSTR_BLOB;
   m_Params[iIndex].pVoid = &str;
   m_Params[iIndex].len = -1;
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

#endif // __ATLSTR_H__

BOOL CSqlite3Command::_Error(int iError, int iNative, LPCSTR pstrMessage)
{
   return m_pDb->_Error(iError, iNative, pstrMessage);
}

