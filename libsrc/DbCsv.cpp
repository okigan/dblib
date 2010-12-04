
#include "stdafx.h"
#include "DbCSV.h"


#ifndef _ASSERTE
   #define _ASSERTE(x)
#endif

#ifndef INT_MAX
   #define INT_MAX 2147483647
#endif


//////////////////////////////////////////////////////////////
// CCsvSystem
//

CCsvSystem::CCsvSystem()
{
}

CCsvSystem::~CCsvSystem()
{
   Terminate();
}

BOOL CCsvSystem::Initialize()
{
   Terminate();
   return TRUE;
}

void CCsvSystem::Terminate()
{
}

IDbDatabase* CCsvSystem::CreateDatabase()
{
   return new CCsvDatabase(this);
}

IDbRecordset* CCsvSystem::CreateRecordset(IDbDatabase* pDb)
{
   return new CCsvRecordset(static_cast<CCsvDatabase*>(pDb));
}

IDbCommand* CCsvSystem::CreateCommand(IDbDatabase* pDb)
{
   return new CCsvCommand(static_cast<CCsvDatabase*>(pDb));
}


//////////////////////////////////////////////////////////////
// CCsvDatabase
//

CCsvDatabase::CCsvDatabase(CCsvSystem* pSystem) : 
   m_pSystem(pSystem),
   m_pstrText(NULL),
   m_pColumns(NULL),
   m_bFixedWidth(false),
   m_cSep(',')
{
   _ASSERTE(m_pSystem);
   m_errs.m_err.m_pDb = this;
}

CCsvDatabase::~CCsvDatabase()
{
   Close();
}

BOOL CCsvDatabase::Open(HWND /*hWnd*/, LPCTSTR pstrConnectionString, LPCTSTR /*pstrUser*/, LPCTSTR /*pstrPassword*/, long /*iType*/)
{
   _ASSERTE(!::IsBadStringPtr(pstrConnectionString,(UINT)-1));

   Close();

   // Store filename
   ::lstrcpy(m_szFilename, pstrConnectionString);

   // Open and read the entire file into memory
   // NOTE: We'll assume the file is ANSI encoded.
   HANDLE hFile = ::CreateFile(pstrConnectionString, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
   if( hFile == INVALID_HANDLE_VALUE ) return _Error((long)::GetLastError(), _T("File open error"));
   DWORD dwSize = ::GetFileSize(hFile, NULL);
   if( dwSize == 0 ) {
      ::CloseHandle(hFile);
      return _Error(1, _T("Empty file?"));
   }
   m_pstrText = (LPSTR) malloc(dwSize + 1);
   if( m_pstrText == NULL ) return _Error(ERROR_OUTOFMEMORY, _T("Out of memory"));
   DWORD dwRead = 0;
   if( !::ReadFile(hFile, m_pstrText, dwSize, &dwRead, NULL) ) {
      _Error(::GetLastError(), _T("File read error"));      
      ::CloseHandle(hFile);
      return FALSE;
   }
   ::CloseHandle(hFile);
   m_pstrText[dwSize] = '\0';

   // Generate row index
   return _BindColumns();
}

void CCsvDatabase::Close()
{
   if( m_pColumns ) delete [] m_pColumns;
   if( m_pstrText ) free(m_pstrText);
   m_pColumns = NULL;
   m_pstrText = NULL;
}

BOOL CCsvDatabase::ExecuteSQL(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/, DWORD* pdwRowsAffected/*=NULL*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));

   CCsvRecordset rec(this);
   if( !rec.Open(pstrSQL, lType, lOptions) ) return FALSE;
   if( pdwRowsAffected ) *pdwRowsAffected = rec.GetRowCount();
   rec.Close();

   return TRUE;
}

BOOL CCsvDatabase::BeginTrans()
{
   _ASSERTE(false);
   return FALSE;
}

BOOL CCsvDatabase::CommitTrans()
{
   _ASSERTE(false);
   return FALSE;
}

BOOL CCsvDatabase::RollbackTrans()
{
   _ASSERTE(false);
   return FALSE;
}

void CCsvDatabase::SetLoginTimeout(long /*lTimeout*/)
{
   _ASSERTE(false);
}

void CCsvDatabase::SetQueryTimeout(long /*lTimeout*/)
{
   _ASSERTE(false);
}

BOOL CCsvDatabase::IsOpen() const
{
   return m_pstrText != NULL;
}

IDbErrors* CCsvDatabase::GetErrors()
{
   return &m_errs;
}

BOOL CCsvDatabase::_Error(long lErrCode, LPCTSTR pstrMessage)
{
   _ASSERTE(!::IsBadStringPtr(pstrMessage,(UINT)-1));
   m_errs.m_nErrors = 1;
   m_errs.m_err.m_lErrCode = lErrCode;
   _tcsncpy(m_errs.m_err.m_szMsg, pstrMessage, (sizeof(m_errs.m_err.m_szMsg)/sizeof(m_errs.m_err.m_szMsg[0])) - 1);
   return FALSE; // Always return FALSE to allow "if( !rc ) return _Error(x)" construct.
}

BOOL CCsvDatabase::_BindColumns()
{
   USES_CONVERSION;
   // Count number of columns
   LPCSTR p = m_pstrText;
   if( *p == ';' ) p++;  // Sometimes column-definition line starts with a ';'-char
   if( *p == '\r' || *p == '\n' || *p == ' ' ) return _Error(1, _T("Junk at start of file"));
   m_nCols = 1;
   m_cSep = ',';
   m_bFixedWidth = false;
   bool bInsideQuote = false;
   bool bWasSpace = false;
   while( *p != '\n' ) {
      // Look for a possible new separator
      if( *p == ';' && m_nCols == 1 ) m_cSep = *p;
      if( *p == '\0' ) return _Error(2, _T("EOF before columns were defined"));
      // So is this a column, then?
      if( !bInsideQuote && *p == m_cSep ) m_nCols++;
      // Skip skip skip
      if( *p == '\"' ) bInsideQuote = !bInsideQuote;
      p++;
      if( *p == '\n' && bInsideQuote ) return _Error(2, _T("Unclosed quotes in field definition"));
   }
   // Create columns array
   m_pColumns = new CCsvColumn[m_nCols];
   // Ready for new run where we populate the columns
   p = m_pstrText;
   if( *p == ';' ) p++;
   bInsideQuote = false;
   bWasSpace = false;
   int iField = 0;
   int iWidth = 0;
   LPCSTR pstrName = p;
   while( *p != '\n' ) {
      if( !bInsideQuote && *p == m_cSep ) {
         // Populate column information
         m_pColumns[iField].iSize = iWidth;
         m_pColumns[iField].lOffset = p - pstrName;
         // A space before the field-separator indicates fixed-width.
         // The "fixed width"-flag is global so we only need to see it once.
         if( bWasSpace ) {
            m_bFixedWidth = true;
            // Trim name as well
            while( iWidth > 0 && pstrName[iWidth - 1] == ' ' ) iWidth--;
         }
         _tcsncpy(m_pColumns[iField].szName, A2CT(pstrName), iWidth);
         // Prepare for next column
         pstrName = ++p;
         iWidth = 0;
         iField++;
      }
      if( *p == '\"' ) bInsideQuote = !bInsideQuote;
      if( bInsideQuote ) m_pColumns[iField].iType = VT_BSTR;
      bWasSpace = (*p == ' ');
      iWidth++;
      p++;
   }
   return TRUE;
}


//////////////////////////////////////////////////////////////
// CCsvRecordset
//

CCsvRecordset::CCsvRecordset(CCsvDatabase* pDb) : 
   m_pDb(pDb),
   m_lCurRow(INT_MAX),
   m_fAttached(false)
{
   _ASSERTE(m_pDb);
   _ASSERTE(m_pDb->IsOpen());
}

CCsvRecordset::~CCsvRecordset()
{
   Close();
}

BOOL CCsvRecordset::Open(LPCTSTR pstrSQL, long lType, long lOptions)
{
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1)); 
   _ASSERTE(::lstrlen(pstrSQL)==0); pstrSQL; // We don't support SQL in this version; pass empty string!!

   if( !m_pDb->IsOpen() ) return FALSE;
   Close();

   m_lType = lType;
   m_lOptions = lOptions;

   // Make room for the field offsets
   long lEmpty = 0;
   m_aFields.RemoveAll();
   for( int i = 0; i < m_pDb->m_nCols; i++ ) m_aFields.Add(lEmpty);

   if( !_BindRows() ) return FALSE;

   return MoveNext();
}

void CCsvRecordset::Close()
{
   m_fAttached = false;
   m_lCurRow = INT_MAX;
}

inline BOOL CCsvRecordset::IsOpen() const
{
   return m_lCurRow < INT_MAX;
}

DWORD CCsvRecordset::GetRowCount() const
{
   _ASSERTE(IsOpen());
   if( !IsOpen() ) return 0;
   return m_aRows.GetSize();
}

BOOL CCsvRecordset::GetField(short iIndex, long& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   long lOffset = m_aFields[iIndex];
   Data = atol(m_pDb->m_pstrText + lOffset);
   return TRUE;
}

BOOL CCsvRecordset::GetField(short iIndex, float& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   long lOffset = m_aFields[iIndex];
   Data = (float) atof(m_pDb->m_pstrText + lOffset);
   return TRUE;
}

BOOL CCsvRecordset::GetField(short iIndex, double& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   long lOffset = m_aFields[iIndex];
   Data = atof(m_pDb->m_pstrText + lOffset);
   return TRUE;
}

BOOL CCsvRecordset::GetField(short iIndex, bool& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   long lOffset = m_aFields[iIndex];
   // Will check for boolean string (left-aligned):
   //   0, "True", true, Yes, yes
   if( m_pDb->m_pColumns[iIndex].iType == VT_BSTR ) Data = toupper(m_pDb->m_pstrText[lOffset]) != 'N' && toupper(m_pDb->m_pstrText[lOffset + 1]) != 'N';
   else Data = atol(m_pDb->m_pstrText + lOffset) != 0;
   return TRUE;
}

BOOL CCsvRecordset::GetField(short iIndex, LPTSTR pData, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   // Extract string; two types of fields can be found: string, "string"
   long lOffset = m_aFields[iIndex];
   int nLength = (int) cchMax; // Need a signed datatype
   if( nLength <= 0 ) return FALSE;
   LPTSTR pOrig = pData;
   LPCSTR p = m_pDb->m_pstrText + lOffset;
   if( *p == '\"' ) {
      p++;  // Skip quote
      while( *p != '\0' && *p != '\"' && *p != '\r' && *p != '\n' && --nLength > 0 ) *pData++ = *p++;
      _ASSERTE(nLength==0 || *p=='\"');
      *pData = '\0';
   }
   else {
      // Right-aligned field?
      while( *p == ' ' ) p++;
      // Make the copy
      char cSep = m_pDb->m_cSep;
      while( *p != '\0' && *p != cSep && *p != '\r' && *p != '\n' && --nLength > 0 ) *pData++ = *p++;
      // Trim fixed-width field?
      while( m_pDb->m_bFixedWidth && pData > pOrig && *(pData - 1) == ' ' ) --pData;
      *pData = '\0';
   }
   return TRUE;
}

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)

BOOL CCsvRecordset::GetField(short iIndex, CString& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   TCHAR szValue[256] = { 0 };
   if( !GetField(iIndex, szValue, 255) ) {
      Data = _T("");
      return FALSE;
   }
   Data = szValue;
   return TRUE;
}

#endif // __ATLSTR_H__

#if defined(_STRING_)

BOOL CCsvRecordset::GetField(short iIndex, std::string& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   TCHAR szValue[256] = { 0 };
   if( !GetField(iIndex, szValue, 255) ) {
      Date = "";
      return FALSE;
   }
   Data = szValue;
   return TRUE;
}

#endif // __ATLSTR_H__

BOOL CCsvRecordset::GetField(short iIndex, SYSTEMTIME& Data)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   TCHAR szValue[100] = { 0 };
   if( !GetField(iIndex, szValue, 99) ) return FALSE;
   // Use the slow, but neat VARIANT functions to convert to a date/time
   USES_CONVERSION;
   DATE d = { 0 };
   ::VarDateFromStr(T2OLE(szValue), 1033, LOCALE_NOUSEROVERRIDE, &d);  // 1033 = US/UK locale
   return SUCCEEDED( ::VariantTimeToSystemTime(d, &Data) );
}

DWORD CCsvRecordset::GetColumnCount() const
{
   _ASSERTE(IsOpen());
   return m_pDb->m_nCols;
}

DWORD CCsvRecordset::GetColumnSize(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return 0;
   return m_pDb->m_bFixedWidth ? m_pDb->m_pColumns[iIndex].iSize : 255;
}

short CCsvRecordset::GetColumnType(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return DB_TYPE_UNKNOWN;
   // NOTE: We can only return the two types that we can recognize in the CSV header.
   //       We don't wish to analyze the contents of each row to determine the real type.
   switch( m_pDb->m_pColumns[iIndex].iType ) {
   case VT_BSTR:
      return DB_TYPE_CHAR;
   case VT_I4:
      return DB_TYPE_INTEGER;
   default:
      _ASSERTE(false);
      return DB_TYPE_UNKNOWN;
   }
}

BOOL CCsvRecordset::GetColumnName(short iIndex, LPTSTR pstrName, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && m_pDb->m_nCols);
   if( iIndex < 0 || iIndex >= m_pDb->m_nCols ) return FALSE;
   ::lstrcpyn( pstrName, m_pDb->m_pColumns[iIndex].szName, cchMax);
   return TRUE;
}

short CCsvRecordset::GetColumnIndex(LPCTSTR pstrName) const
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrName,(UINT)-1));
   // Search for column by name.
   // NOTE: The search is case insensitive!
   for( short i = 0; i < m_pDb->m_nCols; i++ ) {
      if( ::lstrcmp(pstrName, m_pDb->m_pColumns[i].szName) == 0 ) return i;
   }
   return -1;
}

BOOL CCsvRecordset::IsEOF() const
{
   _ASSERTE(IsOpen());
   return m_lCurRow >= m_aRows.GetSize();
}

BOOL CCsvRecordset::MoveNext()
{
   return MoveCursor(1L, 0L);
}

BOOL CCsvRecordset::MovePrev()
{
   return MoveCursor(-1L, 0L);
}

BOOL CCsvRecordset::MoveTop()
{
   return MoveCursor(0L, 0L);
}

BOOL CCsvRecordset::MoveBottom()
{
   return MoveCursor(0L, (long) m_aRows.GetSize() - 1L);
}

BOOL CCsvRecordset::MoveAbs(DWORD dwPos)
{
   return MoveCursor(0L, (long) dwPos);
}

BOOL CCsvRecordset::MoveCursor(long lDiff, long lPos)
{
   _ASSERTE(IsOpen());
   int lNewRow = m_lCurRow;
   if( lDiff == 0 ) lNewRow = lPos; else lNewRow += lDiff;
   if( lNewRow < 0 ) return FALSE;
   if( lNewRow >= m_aRows.GetSize() ) {
      m_lCurRow = m_aRows.GetSize();
      return FALSE;
   }
   if( !_BindFields(lNewRow) ) return FALSE;
   m_lCurRow = lNewRow;
   return TRUE;
}

DWORD CCsvRecordset::GetRowNumber()
{
   _ASSERTE(IsOpen());
   return m_lCurRow;
}

BOOL CCsvRecordset::NextResultset()
{
   _ASSERTE(false);
   return FALSE;
}

BOOL CCsvRecordset::_Error(long lErrCode, LPCTSTR pstrMessage)
{
   _ASSERTE(IsOpen());
   return m_pDb->_Error(lErrCode, pstrMessage);
}

BOOL CCsvRecordset::_BindRows()
{
   LPCSTR p = m_pDb->m_pstrText;
   // Skip header line
   while( *p != '\n' ) p++;
   // Find all rows in the file. Build large array of offsets into each row!
   m_aRows.RemoveAll();
   while( *p != '\0' ) {
      if( *p == '\n' && *(p + 1) != '\0' ) {
         long lOffset = p - m_pDb->m_pstrText + 1;
         m_aRows.Add(lOffset);
      }
      p++;
   }
   // Reset cursor position
   m_lCurRow = -1;
   return TRUE;
}

BOOL CCsvRecordset::_BindFields(long lRow)
{   
   // Find indexes for each field in this row
   // TODO: Optimize when "fixed-width" option is enabled
   LPCSTR p = m_pDb->m_pstrText + m_aRows[lRow];
   long lOffset = p - m_pDb->m_pstrText;
   bool bInsideQuote = false;
   int iFields = 0;
   char cSep = m_pDb->m_cSep;
   m_aFields[iFields++] = lOffset;
   while( *p != '\0' && *p != '\n' ) {
      if( !bInsideQuote && *p == cSep ) {
         long lOffset = p - m_pDb->m_pstrText + 1;
         m_aFields[iFields++] = lOffset;
      }
      else if( *p == '\"' ) {
         bInsideQuote = !bInsideQuote;
      }
      p++;
   }
   _ASSERTE(iFields==m_pDb->m_nCols);
   return iFields == m_pDb->m_nCols;
}


//////////////////////////////////////////////////////////////
// CCsvCommand
//

CCsvCommand::CCsvCommand(CCsvDatabase* pDb) 
   : m_pDb(pDb),
     m_pszSQL(NULL), 
     m_lOptions(0),
     m_dwRows(0)
{
   _ASSERTE(m_pDb);
   _ASSERTE(m_pDb->IsOpen());
}

CCsvCommand::~CCsvCommand()
{
   Close();
}

BOOL CCsvCommand::Create(LPCTSTR pstrSQL, long /*lType = DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/)
{
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));

   Close();

   _ASSERTE(m_pszSQL==NULL);
   m_pszSQL = new TCHAR[ ::lstrlen(pstrSQL) + 1 ];
   if( m_pszSQL == NULL ) return FALSE;
   ::lstrcpy(m_pszSQL, pstrSQL);

   if( (lOptions & DB_OPTION_PREPARE) != 0 ) {
      _ASSERTE(false);
      return FALSE;
   }

   m_dwRows = 0;
   m_lOptions = lOptions;

   return TRUE;
}

BOOL CCsvCommand::Execute(IDbRecordset* pRecordset /*= NULL*/)
{
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(IsOpen());
   if( pRecordset ) {
      return pRecordset->Open(m_pszSQL, DB_OPEN_TYPE_FORWARD_ONLY, m_lOptions);
   }
   else {
      CCsvRecordset rec(m_pDb);
      if( !rec.Open(m_pszSQL, DB_OPEN_TYPE_FORWARD_ONLY, m_lOptions) ) return FALSE;
      m_dwRows = rec.GetRowCount();
      rec.Close();
      return TRUE;
   }
}

void CCsvCommand::Close()
{
   if( m_pszSQL ) delete [] m_pszSQL;
   m_pszSQL = NULL;
}

BOOL CCsvCommand::IsOpen() const
{
   return m_pszSQL != NULL;
}

DWORD CCsvCommand::GetRowCount() const
{
   _ASSERTE(IsOpen());
   if( !IsOpen() ) return 0;
   return m_dwRows;
}

BOOL CCsvCommand::SetParam(short /*iIndex*/, const long* /*pData*/)
{
   _ASSERTE(IsOpen());
   return FALSE;
}

BOOL CCsvCommand::SetParam(short /*iIndex*/, LPCTSTR /*pData*/, UINT /*cchMax = -1*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(false);
   return FALSE;
}

#if defined(_STRING_)

BOOL CCsvCommand::SetParam(short iIndex, std::string& str)
{
   return SetParam(iIndex, str.c_str());
}

#endif // _STRING_

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)

BOOL CCsvCommand::SetParam(short iIndex, CString& str)
{
   return SetParam(iIndex, (LPCTSTR) str);
}

#endif // __ATLSTR_H__

BOOL CCsvCommand::SetParam(short /*iIndex*/, const bool* /*pData*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(false);
   return FALSE;
}

BOOL CCsvCommand::SetParam(short /*iIndex*/, const float* /*pData*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(false);
   return FALSE;
}

BOOL CCsvCommand::SetParam(short /*iIndex*/, const double* /*pData*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(false);
   return FALSE;
}

BOOL CCsvCommand::SetParam(short /*iIndex*/, const SYSTEMTIME* /*pData*/, short /*iType*/)
{
   _ASSERTE(IsOpen());
   _ASSERTE(false);
   return FALSE;
}

BOOL CCsvCommand::_Error(long lErrCode, LPCTSTR pstrMessage)
{
   _ASSERTE(m_pDb);
   return m_pDb->_Error(lErrCode, pstrMessage);
}


//////////////////////////////////////////////////////////////
// CCsvColumn
//

CCsvColumn::CCsvColumn()
{
   iSize = 0;
   iType = VT_I4;
   lOffset = 0;
   ::ZeroMemory(szName, sizeof(szName));
}


//////////////////////////////////////////////////////////////
// CCsvErrors
//

CCsvErrors::CCsvErrors() : 
   m_nErrors(0L)
{
}

long CCsvErrors::GetCount() const
{
   return m_nErrors;
}

void CCsvErrors::Clear()
{
   m_nErrors = 0;
}

IDbError* CCsvErrors::GetError(short iIndex)
{
   _ASSERTE(iIndex>=0 && iIndex<m_nErrors);
   if( iIndex < 0 || iIndex > m_nErrors ) return NULL;
   return &m_err;
}


//////////////////////////////////////////////////////////////
// CCsvError
//

CCsvError::CCsvError()
{
   ::ZeroMemory(m_szMsg, sizeof(m_szMsg));
}

CCsvError::~CCsvError()
{
}

long CCsvError::GetErrorCode()
{
   return m_lErrCode;
}

long CCsvError::GetNativeErrorCode()
{
   return m_lErrCode;
}

void CCsvError::GetOrigin(LPTSTR pstrStr, UINT cchMax)
{
   ::lstrcpyn(pstrStr, _T("CSV-Db"), cchMax);
}

void CCsvError::GetMessage(LPTSTR pstrStr, UINT cchMax)
{
   ::lstrcpyn(pstrStr, m_szMsg, cchMax);
}

void CCsvError::GetSource(LPTSTR pstrStr, UINT cchMax)
{
   _ASSERTE(m_pDb);
   ::lstrcpyn(pstrStr, m_pDb->m_szFilename, cchMax);
}

