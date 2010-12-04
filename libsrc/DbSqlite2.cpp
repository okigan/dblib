
#include "stdafx.h"
#include "DbSqlite2.h"

#ifndef _ASSERTE
   #define _ASSERTE(x)
#endif


//////////////////////////////////////////////////////////////
// CSqlite2System
//

BOOL CSqlite2System::Initialize()
{   
   return TRUE; // No action required!
}

void CSqlite2System::Terminate()
{
}

IDbDatabase*  CSqlite2System::CreateDatabase()
{
   return new CSqlite2Database(this);
}

IDbRecordset* CSqlite2System::CreateRecordset(IDbDatabase* pDb)
{
   return new CSqlite2Recordset(reinterpret_cast<CSqlite2Database*>(pDb));
}

IDbCommand* CSqlite2System::CreateCommand(IDbDatabase* pDb)
{
   return new CSqlite2Command(reinterpret_cast<CSqlite2Database*>(pDb));
}


//////////////////////////////////////////////////////////////
// CSqlite2Error
//

long CSqlite2Error::GetErrorCode()
{
   return 0;
}

long CSqlite2Error::GetNativeErrorCode()
{
   return m_iError;
}

void CSqlite2Error::GetOrigin(LPTSTR pstrStr, UINT cchMax)
{
   _tcsncpy(pstrStr, _T("SQLite"), cchMax);
}

void CSqlite2Error::GetSource(LPTSTR pstrStr, UINT cchMax)
{
   _tcsncpy(pstrStr, _T("SQLite"), cchMax);
}

void CSqlite2Error::GetMessage(LPTSTR pstrStr, UINT cchMax)
{
   USES_CONVERSION;
   _tcsncpy(pstrStr, A2CT(m_szMsg), cchMax);
}

void CSqlite2Error::_Init(int iError, LPCSTR pstrMsg, bool bAutoRelease /*= true*/)
{
   m_iError = iError;
   strncpy(m_szMsg, pstrMsg, sizeof(m_szMsg));
   m_szMsg[ sizeof(m_szMsg) - 1 ] = '\0';
   if( bAutoRelease ) ::sqlite_freemem((LPVOID)pstrMsg);
#if defined(WIN32) && defined(_DEBUG)
   ::MessageBoxA(NULL, m_szMsg, "SQL Error", MB_ICONERROR | MB_SETFOREGROUND);
#endif
}


//////////////////////////////////////////////////////////////
// CSqlite2Errors
//

void CSqlite2Errors::Clear()
{
   m_lCount = 0;
}

long CSqlite2Errors::GetCount() const
{
   return m_lCount;
}

IDbError* CSqlite2Errors::GetError(short iIndex)
{
   if( iIndex != 0 ) return NULL;
   return &m_Error;
}

void CSqlite2Errors::_Init(int iError, LPCSTR pstrMessage, bool bAutoRelease /*= true*/)
{
   m_Error._Init(iError, pstrMessage, bAutoRelease);
}


//////////////////////////////////////////////////////////////
// CSqlite2Database
//

CSqlite2Database::CSqlite2Database(CSqlite2System* pSystem) : m_pSystem(pSystem), m_pDb(NULL)
{
}

CSqlite2Database::~CSqlite2Database()
{
   Close();
}

BOOL CSqlite2Database::Open(HWND /*hWnd*/, LPCTSTR pstrConnectionString, LPCTSTR /*pstrUser*/, LPCTSTR /*pstrPassword*/, long /*iType = DB_OPEN_DEFAULT*/)
{
   _ASSERTE(m_pSystem);
   _ASSERTE(pstrConnectionString);
   if( ::GetFileAttributes(pstrConnectionString) == (DWORD) - 1 ) return _Error(1, "Database does not exists", false);
   Close();
   USES_CONVERSION;
   LPSTR pstrError = NULL;   
   m_pDb = ::sqlite_open(T2CA(pstrConnectionString), 0, &pstrError);
   if( m_pDb == NULL ) return _Error(0, pstrError);
   ::sqlite_busy_timeout(m_pDb, 10000L);
   return TRUE;
}

void CSqlite2Database::Close()
{
   if( m_pDb == NULL ) return;
   ::sqlite_close(m_pDb);
   m_pDb = NULL;
}

BOOL CSqlite2Database::IsOpen() const
{
   return m_pDb != NULL;
}

BOOL CSqlite2Database::ExecuteSQL(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/, DWORD* pdwRowsAffected /*= NULL*/)
{
   CSqlite2Command cmd(this);
   if( !cmd.Create(pstrSQL, lType, lOptions) ) return FALSE;
   if( !cmd.Execute() ) return FALSE;
   if( pdwRowsAffected ) *pdwRowsAffected = cmd.GetRowCount();
   return TRUE;
}

BOOL CSqlite2Database::BeginTrans()
{
   return ExecuteSQL(_T("BEGIN TRANSACTION"));
}

BOOL CSqlite2Database::CommitTrans()
{
   return ExecuteSQL(_T("COMMIT TRANSACTION"));
}

BOOL CSqlite2Database::RollbackTrans()
{
   return ExecuteSQL(_T("ROLLBACK TRANSACTION"));
}

void CSqlite2Database::SetLoginTimeout(long /*lTimeout*/)
{
}

void CSqlite2Database::SetQueryTimeout(long lTimeout)
{
   ::sqlite_busy_timeout(m_pDb, lTimeout);
}

IDbErrors* CSqlite2Database::GetErrors()
{
   return &m_Errors;
}

int CSqlite2Database::GetLastRowId()
{
   return ::sqlite_last_insert_rowid(m_pDb);
}

BOOL CSqlite2Database::_Error(int iError, LPCSTR pstrMessage, bool bAutoRelease /*= true*/)
{   
   m_Errors._Init(iError, pstrMessage, bAutoRelease);
   return FALSE; // Always returns FALSE!
}


//////////////////////////////////////////////////////////////
// CSqlite2Recordset
//

CSqlite2Recordset::CSqlite2Recordset(CSqlite2Database* pDb) : m_pDb(pDb), m_pVm(NULL), m_ppSnapshot(NULL), m_bAttached(false)
{
}

CSqlite2Recordset::~CSqlite2Recordset()
{
   Close();
}

BOOL CSqlite2Recordset::Open(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/)
{
   _ASSERTE(m_pDb);
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));

   Close();

   m_ppSnapshot = NULL;
   m_ppColumns = NULL;
   m_pVm = NULL;
   m_lType = lType;
   m_lOptions = lOptions;
   m_nCols = 0;
   m_nRows = 0;
   m_fEOF = TRUE;

   LPSTR pstrMessage = NULL;
   switch( lType ) {
   case DB_OPEN_TYPE_FORWARD_ONLY:
      {
         LPCSTR pTrail = NULL;
         USES_CONVERSION;
         int iErr = ::sqlite_compile(*m_pDb, T2CA(pstrSQL), &pTrail, &m_pVm, &pstrMessage);
         if( iErr != SQLITE_OK ) return _Error(iErr, pstrMessage);
         _ASSERTE(strlen(pTrail)==0);
         m_iPos = -1;
         m_fEOF = FALSE;
         MoveNext();
         return TRUE;
      }
      break;
   default:
      {
         USES_CONVERSION;
         int iErr = ::sqlite_get_table(*m_pDb, T2CA(pstrSQL), &m_ppSnapshot, &m_nRows, &m_nCols, &pstrMessage);
         if( iErr != SQLITE_OK ) return _Error(iErr, pstrMessage);
         m_iPos = 0;
         return TRUE;
      }
   }
}

void CSqlite2Recordset::Close()
{
   if( m_ppSnapshot && m_lType != DB_OPEN_TYPE_FORWARD_ONLY ) {
      ::sqlite_free_table(m_ppSnapshot);
      m_ppSnapshot = NULL;
   }
   if( m_pVm && !m_bAttached ) {
      ::sqlite_finalize(m_pVm, NULL);
      m_pVm = NULL;
   }
   m_bAttached = false;
}

BOOL CSqlite2Recordset::IsOpen() const
{
   return (m_ppSnapshot != NULL) || m_fEOF;
}

DWORD CSqlite2Recordset::GetRowCount() const
{
   _ASSERTE(IsOpen());
   return ::sqlite_changes(*m_pDb);
}

BOOL CSqlite2Recordset::GetField(short iIndex, long& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   if( pstr == NULL ) {
      Data = 0;
      return TRUE;
   }
   else {
      Data = atol(pstr);
   }
   return TRUE;
}

BOOL CSqlite2Recordset::GetField(short iIndex, float& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   if( pstr == NULL ) {
      Data = 0.0f;
      return TRUE;
   }
   else {
      Data = (float) atof(pstr);
   }
   return TRUE;
}

BOOL CSqlite2Recordset::GetField(short iIndex, double& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   if( pstr == NULL ) {
      Data = 0.0;
      return TRUE;
   }
   else {
      Data = atof(pstr);
   }
   return TRUE;
}

BOOL CSqlite2Recordset::GetField(short iIndex, bool& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   if( pstr == NULL ) {
      Data = false;
   }
   else {
      // 1, -1, True, true, Yes, yes
      Data = *pstr == '1' || *pstr == '-' || *pstr == 'T' || *pstr == 't' || *pstr == 'Y' || *pstr == 'y';
   }
   return TRUE;
}

BOOL CSqlite2Recordset::GetField(short iIndex, LPTSTR pData, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   if( pstr == NULL ) {
      _tcscpy(pData, _T(""));
   }
   else {
      USES_CONVERSION;
      LPCSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
      _tcsncpy(pData, A2CT(pstr), cchMax);
   }
   return TRUE;
}

BOOL CSqlite2Recordset::GetField(short iIndex, SYSTEMTIME& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   ::ZeroMemory(&Data, sizeof(SYSTEMTIME));
   LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   if( pstr == NULL ) return TRUE;
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

BOOL CSqlite2Recordset::GetField(short iIndex, CString& pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   if( pstr == NULL ) {
      pData = _T("");
   }
   else {
      pData = pstr;
   }
   return TRUE;
}

#endif // __ATLSTR_H__

#if defined(_STRING_)

BOOL CSqlite2Recordset::GetField(short iIndex, std::string& pData)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( IsEOF() ) return FALSE;
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   LPSTR pstr = m_ppSnapshot[ ((m_iPos + 1) * m_nCols) + iIndex ];
   if( pstr == NULL ) {
      pData = "";
   }
   else {
      pData = pstr;
   }
   return TRUE;
}

#endif // __ATLSTR_H__

DWORD CSqlite2Recordset::GetColumnSize(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex < 0 || iIndex >= m_nCols ) return 0;
   return 255;  // Hardcoded value! Sqlite actually supports up to many Mb of data!
}

BOOL CSqlite2Recordset::GetColumnName(short iIndex, LPTSTR pstrName, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   USES_CONVERSION;
   LPCSTR pstrSrc = m_lType == DB_OPEN_TYPE_FORWARD_ONLY ? m_ppColumns[iIndex] : m_ppSnapshot[iIndex];
   _tcsncpy(pstrName, A2CT(pstrSrc), cchMax);
   return TRUE;
}

short CSqlite2Recordset::GetColumnType(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex < 0 || iIndex >= m_nCols ) return FALSE;
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      LPCSTR pstr = m_ppColumns[iIndex + m_nCols];
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
   else {
      _ASSERTE(false); // Not supported for this recordset type
      return DB_TYPE_UNKNOWN;
   }
}

short CSqlite2Recordset::GetColumnIndex(LPCTSTR pstrName) const
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrName,(UINT)-1));
   USES_CONVERSION;
   LPCSTR pstr = T2CA(pstrName);
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      for( short i = 0; i < m_nCols; i++ ) if( strcmp(pstr, m_ppColumns[i]) == 0 ) return i;
   }
   else {
      for( short i = 0; i < m_nCols; i++ ) if( strcmp(pstr, m_ppSnapshot[i]) == 0 ) return i;
   }
   return -1;
}

DWORD CSqlite2Recordset::GetColumnCount() const
{
   _ASSERTE(IsOpen());
   return m_nCols;
}

BOOL CSqlite2Recordset::IsEOF() const
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return m_fEOF;
   return m_iPos >= m_nRows;
}

BOOL CSqlite2Recordset::MoveNext()
{   
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      int iErr = ::sqlite_step(m_pVm, &m_nCols, const_cast<const char***>(&m_ppSnapshot), &m_ppColumns);
      m_fEOF = iErr != SQLITE_ROW;
      _ASSERTE(iErr==SQLITE_DONE || iErr==SQLITE_ROW);
      if( iErr != SQLITE_DONE && iErr != SQLITE_ROW ) return _Error(1, "Move Error", false);
      return TRUE;
   }
   else {
      ATLASSERT(IsOpen());
      m_iPos++;
      return IsEOF();
   }
}

BOOL CSqlite2Recordset::MovePrev()
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return _Error(1, "Invalid recordset type", false);
   if( m_iPos <= 0 ) return FALSE;
   --m_iPos;
   return TRUE;
}

BOOL CSqlite2Recordset::MoveTop()
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return _Error(1, "Invalid recordset type", false);
   m_iPos = 0;
   return TRUE;
}

BOOL CSqlite2Recordset::MoveBottom()
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return _Error(1, "Invalid recordset type", false);
   if( m_nRows == 0 ) return FALSE;
   m_iPos = m_nRows - 1;
   return IsEOF();
}

BOOL CSqlite2Recordset::MoveAbs(DWORD dwPos)
{
   _ASSERTE(IsOpen());
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return _Error(1, "Invalid recordset type", false);
   if( m_iPos < 0 ) return FALSE;
   if( m_iPos >= m_nRows ) return FALSE;
   m_iPos = (int) dwPos;
   return TRUE;
}

DWORD CSqlite2Recordset::GetRowNumber()
{
   _ASSERTE(IsOpen());
   _ASSERT(m_lType!=DB_OPEN_TYPE_FORWARD_ONLY);
   if( m_lType == DB_OPEN_TYPE_FORWARD_ONLY ) return 0;
   return m_iPos;
}

BOOL CSqlite2Recordset::NextResultset()
{
   _ASSERTE(IsOpen());
   _ASSERTE(false);   // Not supported
   return _Error(1, "Unsupported", false);
}

BOOL CSqlite2Recordset::_Error(int iError, LPCSTR pstrMessage, bool bAutoRelease /*= true*/)
{
   m_pDb->_Error(iError, pstrMessage, bAutoRelease);
   return FALSE;
}

BOOL CSqlite2Recordset::_Attach(sqlite_vm* pVm)
{
   Close();
   m_pVm = pVm;
   m_ppSnapshot = NULL;
   m_lType = DB_OPEN_TYPE_FORWARD_ONLY;
   m_lOptions = 0;
   m_nCols = 0;
   m_nRows = 0;
   m_iPos = -1;
   m_fEOF = FALSE;
   m_bAttached = true;
   return MoveNext();
}


//////////////////////////////////////////////////////////////
// CSqlite2Command
//

CSqlite2Command::CSqlite2Command(CSqlite2Database* pDb) : m_pDb(pDb), m_pVm(NULL), m_nParams(0)
{
}

CSqlite2Command::~CSqlite2Command()
{
   Close();
}

BOOL CSqlite2Command::Create(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long /*lOptions = DB_OPTION_DEFAULT*/)
{
   _ASSERTE(m_pDb);
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));
   _ASSERTE(lType==DB_OPEN_TYPE_FORWARD_ONLY);

   Close();

   if( lType != DB_OPEN_TYPE_FORWARD_ONLY ) return FALSE;

   LPCSTR pTrail = NULL;
   LPSTR pstrError = NULL;
   USES_CONVERSION;
   int iErr = ::sqlite_compile(*m_pDb, T2CA(pstrSQL), &pTrail, &m_pVm, &pstrError);
   if( iErr != SQLITE_OK ) return _Error(iErr, pstrError);

   m_nParams = 0;

   return TRUE;
}

BOOL CSqlite2Command::Execute(IDbRecordset* pRecordset /*= NULL*/)
{
   USES_CONVERSION;

   LPSTR pstrError = NULL;
   int iErr = ::sqlite_reset(m_pVm, &pstrError);
   if( iErr != SQLITE_OK ) return _Error(iErr, pstrError);

   LPCSTR pstr = NULL;
   CHAR szBuffer[32];
   int iLen;
   for( int iIndex = 0; iIndex < m_nParams; iIndex++ ) {
      TSqlite2Param& Param = m_Params[iIndex];
      switch( Param.vt ) {
      case VT_I4:
         iLen = ::wsprintfA(szBuffer, "%ld", * (char*) Param.pVoid);
         iErr = ::sqlite_bind(m_pVm, iIndex + 1, szBuffer, iLen + 1, TRUE);
         break;
      case VT_R4:
         iLen = ::wsprintfA(szBuffer, "%f", (double) * (float*) Param.pVoid);
         iErr = ::sqlite_bind(m_pVm, iIndex + 1, szBuffer, iLen + 1, TRUE);
         break;
      case VT_R8:
         iLen = ::wsprintfA(szBuffer, "%f", * (long*) Param.pVoid);
         iErr = ::sqlite_bind(m_pVm, iIndex + 1, szBuffer, iLen + 1, TRUE);
         break;
      case VT_LPSTR:
         pstr = T2A( (LPTSTR) Param.pVoid );
         if( Param.len == -1 ) iErr = ::sqlite_bind(m_pVm, iIndex + 1, pstr, -1, TRUE);
         else if( Param.len == -2 ) iErr = ::sqlite_bind(m_pVm, iIndex + 1, NULL, 0, TRUE);
         else iErr = ::sqlite_bind(m_pVm, iIndex + 1, pstr, Param.len + 1, TRUE);
         break;
      case VT_BOOL:
         strcpy(szBuffer, * (bool*) Param.pVoid ? "1" : "0");
         iErr = ::sqlite_bind(m_pVm, iIndex + 1, szBuffer, 2, TRUE);
         break;
#if defined(_STRING_)
      case VT_BSTR:
         pstr = (const char*) ((std::string*) Param.pVoid)->c_str();
         iErr = ::sqlite_bind(m_pVm, iIndex + 1, pstr, strlen(pstr) + 1, TRUE);
         break;
#endif
#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)
      case VT_BSTR_BLOB:
         pstr = T2A( * ((CString*) Param.pVoid) );
         iErr = ::sqlite_bind(m_pVm, iIndex + 1, pstr, strlen(pstr) + 1, TRUE);
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
            if( pData->wYear == 0 ) iErr = ::sqlite_bind(m_pVm, iIndex + 1, NULL, 0, TRUE);
            else iErr = ::sqlite_bind(m_pVm, iIndex + 1, szBuffer, strlen(szBuffer) + 1, TRUE);
         }
         break;
      }
      if( iErr != SQLITE_OK ) return _Error(iErr, "Bind Error", false);
   }

   if( pRecordset ) {
      CSqlite2Recordset* pRec = reinterpret_cast<CSqlite2Recordset*>(pRecordset);
      return pRec->_Attach(m_pVm);
   }
   else {
      CSqlite2Recordset rec(m_pDb);
      return rec._Attach(m_pVm);
   }
}

void CSqlite2Command::Close()
{
   if( m_pVm ) {
      ::sqlite_finalize(m_pVm, NULL);
      m_pVm = NULL;
   }
}

BOOL CSqlite2Command::IsOpen() const
{
   return m_pVm != NULL;
}

DWORD CSqlite2Command::GetRowCount() const
{
   _ASSERTE(IsOpen());
   return ::sqlite_changes(*m_pDb);
}

BOOL CSqlite2Command::SetParam(short iIndex, const long* pData)
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

BOOL CSqlite2Command::SetParam(short iIndex, LPCTSTR pData, UINT cchMax /*= -1*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex<DB_MAX_PARAMS);
   if( iIndex < 0 || iIndex >= DB_MAX_PARAMS ) return FALSE;
   m_Params[iIndex].vt = VT_LPSTR;
   m_Params[iIndex].pVoid = pData;
   m_Params[iIndex].len = (int) cchMax;
   if( ++iIndex > m_nParams ) m_nParams = iIndex;
   return TRUE;
}

BOOL CSqlite2Command::SetParam(short iIndex, const bool* pData)
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

BOOL CSqlite2Command::SetParam(short iIndex, const float* pData)
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

BOOL CSqlite2Command::SetParam(short iIndex, const double* pData)
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

BOOL CSqlite2Command::SetParam(short iIndex, const SYSTEMTIME* pData, short iType /*= DB_TYPE_TIMESTAMP*/)
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

BOOL CSqlite2Command::SetParam(short iIndex, std::string& str)
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

BOOL CSqlite2Command::SetParam(short iIndex, CString& str)
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

BOOL CSqlite2Command::_Error(int iError, LPCSTR pstrMessage, bool bAutoRelease /*= true*/)
{
   return m_pDb->_Error(iError, pstrMessage, bAutoRelease);
}

