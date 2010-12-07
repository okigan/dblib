
#include "stdafx.h"
#include "dblib/DbOledb.h"

#if _MSC_VER < 1300
   #pragma warning(disable : 4244)
#endif


//////////////////////////////////////////////////////////////
// COledbSystem
//

COledbSystem::COledbSystem()
{
}

COledbSystem::~COledbSystem()
{
}

BOOL COledbSystem::Initialize()
{
   ::CoInitialize(NULL);
   return TRUE;
}

void COledbSystem::Terminate()
{
   ::CoUninitialize();
}

IDbDatabase* COledbSystem::CreateDatabase()
{
   return new COledbDatabase(this);
}

IDbRecordset* COledbSystem::CreateRecordset(IDbDatabase* pDb)
{
   return new COledbRecordset(static_cast<COledbDatabase*>(pDb));
}

IDbCommand* COledbSystem::CreateCommand(IDbDatabase* pDb)
{
   return new COledbCommand(static_cast<COledbDatabase*>(pDb));
}


//////////////////////////////////////////////////////////////
// COledbDatabase
//

COledbDatabase::COledbDatabase(COledbSystem* pSystem) : 
   m_pSystem(pSystem)
{
   _ASSERTE(pSystem);
   ATLTRY(m_pErrors = new COledbErrors);
#ifdef _DEBUG
   m_nRecordsets = 0;
#endif
}

COledbDatabase::~COledbDatabase()
{
   Close();
   delete m_pErrors;
#ifdef _DEBUG
   // Check that all recordsets have been closed and deleted
   _ASSERTE(m_nRecordsets==0);
#endif
}

BOOL COledbDatabase::Open(HWND hWnd, LPCTSTR pstrConnectionString, LPCTSTR pstrUser, LPCTSTR pstrPassword, long iType)
{
   HRESULT Hr;

   Close();

   if( (iType & DB_OPEN_PROMPTDIALOG) != 0 ) 
   {
      // Initialize from Data Links UI

      CComPtr<IDBPromptInitialize> spPrompt;
      Hr = spPrompt.CoCreateInstance(CLSID_DataLinks);
      if( FAILED(Hr) ) return _Error(Hr);

      if( hWnd == NULL ) hWnd = ::GetActiveWindow();
      if( hWnd == NULL ) hWnd = ::GetDesktopWindow();

      Hr = spPrompt->PromptDataSource(
         NULL,                             // pUnkOuter
         hWnd,                             // hWndParent
         DBPROMPTOPTIONS_PROPERTYSHEET,    // dwPromptOptions
         0,                                // cSourceTypeFilter
         NULL,                             // rgSourceTypeFilter
         NULL,                             // pwszszzProviderFilter
         IID_IDBInitialize,                // riid
         (LPUNKNOWN*) &m_spInit);          // ppDataSource
      if( Hr == S_FALSE ) Hr = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_WIN32, ERROR_CANCELLED);  // The user clicked cancel
      if( FAILED(Hr) ) return _Error(Hr);
   }
   else 
   {
      TCHAR szPrefix[] = _T("Provider=");
      ::lstrcpyn(szPrefix, pstrConnectionString, (sizeof(szPrefix) / sizeof(TCHAR)));
      if( ::lstrcmpi(szPrefix, _T("Provider=")) == 0 ) 
      {
         // Connect using OLE DB connection string

         CComPtr<IDataInitialize> spDataInit;
         Hr = spDataInit.CoCreateInstance(CLSID_MSDAINITIALIZE);
         if( FAILED(Hr) ) return _Error(Hr);

         USES_CONVERSION;
         Hr = spDataInit->GetDataSource(NULL, CLSCTX_INPROC_SERVER, T2OLE((LPTSTR)pstrConnectionString), IID_IDBInitialize, (LPUNKNOWN*)&m_spInit);
         if( FAILED(Hr) ) return _Error(Hr);
      }
      else 
      {
         // Initialize from ODBC DSN information

         Hr = m_spInit.CoCreateInstance(L"MSDASQL");
         if( FAILED(Hr) ) return _Error(Hr);

         DBPROPSET PropSet;
         const ULONG nMaxProps = 4;
         DBPROP Prop[nMaxProps];
         ULONG iProp = 0;

         // Initialize common property options.
         ULONG i;
         for( i = 0; i < nMaxProps; i++ ) {
            ::VariantInit(&Prop[i].vValue);
            Prop[i].dwOptions = DBPROPOPTIONS_REQUIRED;
            Prop[i].colid = DB_NULLID;
         }

         // Level of prompting that will be done to complete the connection process
         Prop[iProp].dwPropertyID = DBPROP_INIT_PROMPT;
         Prop[iProp].vValue.vt = VT_I2;
         Prop[iProp].vValue.iVal = DBPROMPT_NOPROMPT;    
         iProp++;
         // Data source name--see the sample source included with the OLE DB SDK.
         Prop[iProp].dwPropertyID = DBPROP_INIT_DATASOURCE;   
         Prop[iProp].vValue.vt = VT_BSTR;
         Prop[iProp].vValue.bstrVal = T2BSTR(pstrConnectionString);
         iProp++;
         if( pstrUser ) {
            // User ID
            Prop[iProp].dwPropertyID = DBPROP_AUTH_USERID;
            Prop[iProp].vValue.vt = VT_BSTR;
            Prop[iProp].vValue.bstrVal = T2BSTR(pstrUser);
            iProp++;
         }
         if( pstrPassword ) {
            // Password
            Prop[iProp].dwPropertyID = DBPROP_AUTH_PASSWORD;
            Prop[iProp].vValue.vt = VT_BSTR;
            Prop[iProp].vValue.bstrVal = T2BSTR(pstrPassword);
            iProp++;
         }

         // Prepare properties
         PropSet.guidPropertySet = DBPROPSET_DBINIT;
         PropSet.cProperties = iProp;
         PropSet.rgProperties = Prop;
         // Set initialization properties.
         CComQIPtr<IDBProperties> spProperties = m_spInit;
         Hr = spProperties->SetProperties(1, &PropSet);
      
         // Before we check if it failed, clean up
         for( i = 0; i < nMaxProps; i++ ) ::VariantClear(&Prop[i].vValue);
      
         // Did SetProperties() fail?
         if( FAILED(Hr) ) return _Error(Hr);
      }
   }

   return Connect();
}

BOOL COledbDatabase::Connect()
{
   _ASSERTE(m_spInit);
   if( m_spInit == NULL ) return FALSE;

   // Initialize datasource
   HRESULT Hr = m_spInit->Initialize();
   if( FAILED(Hr) ) return _Error(Hr);

   // Create session
   CComQIPtr<IDBCreateSession> spCreateSession = m_spInit;
   if( spCreateSession == NULL ) return FALSE;
   Hr = spCreateSession->CreateSession(NULL, IID_IOpenRowset, (LPUNKNOWN*) &m_spSession);
   if( FAILED(Hr) ) return _Error(Hr);

   return TRUE;
}

void COledbDatabase::Close()
{
   m_spSession.Release();
   if( m_spInit ) {
      m_spInit->Uninitialize();
      m_spInit.Release();
   }
}

BOOL COledbDatabase::IsOpen() const
{
   return m_spSession != NULL;
}

BOOL COledbDatabase::ExecuteSQL(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/, DWORD* pdwRowsAffected/*= NULL*/)
{
   USES_CONVERSION;
   HRESULT Hr;

   if( pdwRowsAffected ) *pdwRowsAffected = 0;
   CComQIPtr<IDBCreateCommand> spCreate = m_spSession;
   if( spCreate == NULL ) return FALSE;
   CComPtr<ICommand> spCommand;
   Hr = spCreate->CreateCommand(NULL, IID_ICommand, (LPUNKNOWN*) &spCommand);
   if( FAILED(Hr) ) return _Error(Hr);
   _SetRecordsetType(spCommand, lType, lOptions);
   CComQIPtr<ICommandText> spText = spCommand;
   _ASSERTE(spText);
   Hr = spText->SetCommandText(DBGUID_DBSQL, T2CW(pstrSQL));
   if( FAILED(Hr) ) return _Error(Hr);
   CComPtr<IRowset> spRowset;
   Hr = spText->Execute(NULL, IID_IRowset, NULL, (LONG*) pdwRowsAffected, (LPUNKNOWN*) &spRowset);
   if( FAILED(Hr) ) return _Error(Hr);
   return TRUE;
}

BOOL COledbDatabase::BeginTrans()
{
   return BeginTrans(ISOLATIONLEVEL_READCOMMITTED);
}

BOOL COledbDatabase::BeginTrans(ISOLEVEL isoLevel)
{
   CComQIPtr<ITransactionLocal> spTransaction = m_spSession;
   if( spTransaction == NULL ) return FALSE;
   HRESULT Hr = spTransaction->StartTransaction(isoLevel, 0, NULL, NULL);
   return SUCCEEDED(Hr) ? TRUE : _Error(Hr);
}

BOOL COledbDatabase::CommitTrans()
{
   CComQIPtr<ITransaction> spTransaction = m_spSession;
   if( spTransaction == NULL ) return FALSE;
   BOOL bRetaining = FALSE;
   DWORD grfTC = XACTTC_SYNC;
   DWORD grfRM = 0;
   HRESULT Hr = spTransaction->Commit(bRetaining, grfTC, grfRM);
   _ASSERTE(SUCCEEDED(Hr));
   return SUCCEEDED(Hr) ? TRUE : _Error(Hr);
}

BOOL COledbDatabase::RollbackTrans()
{
   CComQIPtr<ITransaction> spTransaction = m_spSession;
   if( spTransaction == NULL ) return FALSE;
   BOID* pboidReason = NULL;
   BOOL bRetaining = FALSE;
   BOOL bAsync = FALSE;
   HRESULT Hr = spTransaction->Abort(pboidReason, bRetaining, bAsync);
   _ASSERTE(SUCCEEDED(Hr));
   return SUCCEEDED(Hr) ? TRUE : _Error(Hr);
}

void COledbDatabase::SetLoginTimeout(long lTimeout)
{
   m_lLoginTimeout = lTimeout;
}

void COledbDatabase::SetQueryTimeout(long lTimeout)
{
   m_lQueryTimeout = lTimeout;
}

IDbErrors* COledbDatabase::GetErrors()
{
   _ASSERTE(m_pErrors);
   return m_pErrors;
}

BOOL COledbDatabase::_Error(HRESULT Hr)
{
   _ASSERTE(m_pErrors);
   m_pErrors->_Init(Hr);
   return FALSE; // Always return FALSE to allow "return _Error();" constructs...  
}

void COledbDatabase::_SetRecordsetType(ICommand* pCommand, long lType, long /*lOptions*/)
{
   DBPROPSET PropSet;
   const ULONG nMaxProps = 8;
   DBPROP Prop[nMaxProps];
   ULONG iProp = 0;

   // Initialize common property options.
   ULONG i;
   for( i = 0; i < nMaxProps; i++ ) {
      ::VariantInit(&Prop[i].vValue);
      Prop[i].dwOptions = DBPROPOPTIONS_REQUIRED;
      Prop[i].colid = DB_NULLID;
      Prop[1].dwStatus = 0;
   }

   if( lType == DB_OPEN_TYPE_FORWARD_ONLY ) {
      // None
   }
   else if( lType == DB_OPEN_TYPE_SNAPSHOT ) {
      Prop[iProp].dwPropertyID = DBPROP_CANFETCHBACKWARDS;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_TRUE;
      iProp++;
      Prop[iProp].dwPropertyID = DBPROP_CANSCROLLBACKWARDS;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_TRUE;
      iProp++;
/*
      Prop[iProp].dwPropertyID = DBPROP_CANHOLDROWS;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_TRUE;
      iProp++;
*/
   }
   else if( lType == DB_OPEN_TYPE_DYNASET ) {
      Prop[iProp].dwPropertyID = DBPROP_BOOKMARKS;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_FALSE;
      iProp++;
      Prop[iProp].dwPropertyID = DBPROP_CANFETCHBACKWARDS;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_TRUE;
      iProp++;
      Prop[iProp].dwPropertyID = DBPROP_CANSCROLLBACKWARDS;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_TRUE;
      iProp++;
      Prop[iProp].dwPropertyID = DBPROP_OTHERINSERT;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_TRUE;
      iProp++;
      Prop[iProp].dwPropertyID = DBPROP_OTHERUPDATEDELETE;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_TRUE;
      iProp++;
      Prop[iProp].dwPropertyID = DBPROP_REMOVEDELETED;
      Prop[iProp].vValue.vt = VT_BOOL;
      Prop[iProp].vValue.boolVal = VARIANT_TRUE;
      iProp++;
   }
   else {
      _ASSERTE(false);
   }

   // Prepare properties
   PropSet.guidPropertySet = DBPROPSET_ROWSET;
   PropSet.cProperties = iProp;
   PropSet.rgProperties = Prop;
   // Set initialization properties.
   CComQIPtr<ICommandProperties> spProperties = pCommand;
   if( iProp > 0 ) {
      HRESULT Hr = spProperties->SetProperties(1, &PropSet);
      _ASSERTE(SUCCEEDED(Hr));
      Hr;
   }

   // Clean up
   for( i = 0; i < nMaxProps; i++ ) ::VariantClear(&Prop[i].vValue);
}

//////////////////////////////////////////////////////////////
// COledbRecordset
//

COledbRecordset::COledbRecordset(COledbDatabase* pDb) : 
   m_pDb(pDb), 
   m_pData(NULL),
   m_rgBindings(NULL),
   m_hAccessor(DB_NULL_HACCESSOR), 
   m_pwstrNameBuffer(NULL),
   m_fEOF(true)
{
   _ASSERTE(m_pDb==NULL || m_pDb->IsOpen());
#ifdef _DEBUG
   if( m_pDb ) m_pDb->m_nRecordsets++;
#endif
}

COledbRecordset::COledbRecordset(IRowset* pRS) : 
   m_pDb(NULL), 
   m_pData(NULL),
   m_rgBindings(NULL), 
   m_hAccessor(DB_NULL_HACCESSOR), 
   m_pwstrNameBuffer(NULL),
   m_fEOF(true)
{
   Attach(pRS);
}

COledbRecordset::~COledbRecordset()
{
   Close();
#ifdef _DEBUG
   if( m_pDb ) m_pDb->m_nRecordsets--;
#endif
}

BOOL COledbRecordset::Open(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/)
{
   _ASSERTE(m_pDb==NULL || m_pDb->IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));
   HRESULT Hr;

   // Close old recordset
   Close();

   m_nRowsAffected = 0;

   // Create a new recordset
   CComQIPtr<IDBCreateCommand> spCreate = m_pDb->m_spSession;
   if( spCreate == NULL ) return FALSE;
   CComPtr<ICommand> spCommand;
   Hr = spCreate->CreateCommand(NULL, IID_ICommand, (LPUNKNOWN*) &spCommand);
   if( FAILED(Hr) ) return _Error(Hr);
   // Set type
   COledbDatabase::_SetRecordsetType(spCommand, lType, lOptions);
   // Set SQL
   CComQIPtr<ICommandText> spText = spCommand;
   _ASSERTE(spText);
   USES_CONVERSION;
   Hr = spText->SetCommandText(DBGUID_DBSQL, T2COLE(pstrSQL));
   if( FAILED(Hr) ) return _Error(Hr);

   // Execute...
   Hr = spText->Execute(NULL, IID_IRowset, NULL, &m_nRowsAffected, (LPUNKNOWN*) &m_spRowset);
   if( FAILED(Hr) ) return _Error(Hr);

   // Bind columns
   if( !_BindColumns() ) return FALSE;

   return MoveNext();
}

void COledbRecordset::Close()
{
   if( m_pData ) {
      ::CoTaskMemFree(m_pData);
      m_pData = NULL;
   }
   if( m_rgBindings ) {
      ::CoTaskMemFree(m_rgBindings);
      m_rgBindings = NULL;
   }
   if( m_pwstrNameBuffer ) {
      ::CoTaskMemFree(m_pwstrNameBuffer);
      m_pwstrNameBuffer = NULL;
   }
   m_hAccessor = DB_NULL_HACCESSOR;
   m_nCols = 0;
   m_spRowset.Release();
}

BOOL COledbRecordset::IsOpen() const
{
   return m_spRowset != NULL;
}

DWORD COledbRecordset::GetRowCount() const
{
   return m_nRowsAffected;
}

BOOL COledbRecordset::GetField(short iIndex, long& Data)
{
   _ASSERTE(IsOpen());
   iIndex = iIndex + m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_I1:
   case DBTYPE_UI1:
      Data = * (char*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
      return TRUE;
   case DBTYPE_I2:
   case DBTYPE_UI2:
      Data = * (SHORT*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
      return TRUE;
   case DBTYPE_I4:
   case DBTYPE_UI4:
      Data = * (LONG*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COledbRecordset::GetField(short iIndex, float& Data)
{
   _ASSERTE(IsOpen());
   iIndex = iIndex + m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_STR:
      Data = (float) atof( (LPSTR) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue) );
      return TRUE;
   case DBTYPE_R4:
      Data = * (float*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COledbRecordset::GetField(short iIndex, double& Data)
{
   _ASSERTE(IsOpen());
   iIndex = iIndex + m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_STR:
      Data = atof( (LPSTR) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue) );
      return TRUE;
   case DBTYPE_R8:
      Data = * (double*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COledbRecordset::GetField(short iIndex, bool& Data)
{
   _ASSERTE(IsOpen());
   iIndex = iIndex + m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_BOOL:
      // TODO: Verify size of bool
      Data = * (bool*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COledbRecordset::GetField(short iIndex, TCHAR* pData, UINT cchMax)
{
   _ASSERTE(IsOpen());
   iIndex = iIndex + m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   // NOTE: We should never be forced to convert between MBCS/UNICODE because
   //       the column binding sets up the desired TCHAR type.
   USES_CONVERSION;
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_STR:
      ::lstrcpyn(pData, A2CT( (char*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue) ), cchMax);
      return TRUE;
   case DBTYPE_WSTR:
      ::lstrcpyn(pData, W2CT( (WCHAR*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue) ), cchMax);
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

BOOL COledbRecordset::GetField(short iIndex, SYSTEMTIME& Data)
{
   _ASSERTE(IsOpen());
   iIndex = iIndex + m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   // NOTE: Not tested
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_DATE:
      {
         DATE date = * (DATE*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
         ::VariantTimeToSystemTime(date, &Data);
      }
      return TRUE;
   case DBTYPE_DBDATE:
      {
         DBDATE date = * (DBDATE*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
         Data.wYear = date.year;
         Data.wMonth = date.month;
         Data.wDay = date.day;
      }
      return TRUE;
   case DBTYPE_DBTIME:
      {
         DBTIME time = * (DBTIME*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
         Data.wHour = time.hour;
         Data.wMinute = time.minute;
         Data.wSecond = time.second;
      }
      return TRUE;
   case DBTYPE_DBTIMESTAMP:
      {
         DBTIMESTAMP ts = * (DBTIMESTAMP*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
         Data.wYear = ts.year;
         Data.wMonth = ts.month;
         Data.wDay = ts.day;
         Data.wHour = ts.hour;
         Data.wMinute = ts.minute;
         Data.wSecond = ts.second;
      }
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)

BOOL COledbRecordset::GetField(short iIndex, CString& pData)
{
   _ASSERTE(IsOpen());
   iIndex += m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   // NOTE: We should never be forced to convert between MBCS/UNICODE because
   //       the column binding sets up the desired TCHAR type.
   USES_CONVERSION;
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_STR:
      pData = A2CT( (char*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue) );
      return TRUE;
   case DBTYPE_WSTR:
      pData = W2CT( (WCHAR*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue) );
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

#endif // __ATLSTR_H__

#if defined(_STRING_)

BOOL COledbRecordset::GetField(short iIndex, std::string& pData)
{
   _ASSERTE(IsOpen());
   iIndex += m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   USES_CONVERSION;
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_STR:
      pData = (char*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue);
      return TRUE;
   case DBTYPE_WSTR:
      pData = W2CA( (WCHAR*) ((LPBYTE) m_pData + m_rgBindings[iIndex].obValue) );
      return TRUE;
   default:
      _ASSERTE(false);
      return FALSE;
   }
}

#endif // __ATLSTR_H__

DWORD COledbRecordset::GetColumnCount() const
{
   if( !IsOpen() ) return 0;
   return m_nCols - m_iAdjustIndex;
}

DWORD COledbRecordset::GetColumnSize(short iIndex)
{
   _ASSERTE(IsOpen());
   return m_rgBindings[iIndex + m_iAdjustIndex].cbMaxLen;
}

short COledbRecordset::GetColumnType(short iIndex)
{
   _ASSERTE(IsOpen());
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   iIndex = iIndex + m_iAdjustIndex;
   switch( m_rgBindings[iIndex].wType ) {
   case DBTYPE_STR:
   case DBTYPE_WSTR:
      return DB_TYPE_CHAR;
   case DBTYPE_I1:
   case DBTYPE_UI1:
   case DBTYPE_I2:
   case DBTYPE_UI2:
   case DBTYPE_I4:
   case DBTYPE_UI4:
      return DB_TYPE_INTEGER;
   case DBTYPE_BOOL:
      return DB_TYPE_BOOLEAN;
   case DBTYPE_DATE:
   case DBTYPE_DBDATE:
      return DB_TYPE_DATE;
   case DBTYPE_DBTIME:
      return DB_TYPE_TIME;
   case DBTYPE_DBTIMESTAMP:
      return DB_TYPE_TIMESTAMP;
   case DBTYPE_R4:
      return DB_TYPE_REAL;
   case DBTYPE_DECIMAL:
   case DBTYPE_NUMERIC:
   case DBTYPE_R8:
      return DB_TYPE_DOUBLE;
   default:
      return DB_TYPE_UNKNOWN;
   }
}

BOOL COledbRecordset::GetColumnName(short iIndex, TCHAR* pstrName, UINT cchMax)
{
   _ASSERTE(IsOpen());
   _ASSERTE(cchMax>0);
   _ASSERTE(!::IsBadWritePtr(pstrName, cchMax));
   iIndex = iIndex + m_iAdjustIndex;
   _ASSERTE(iIndex>=0 && iIndex<m_nCols);
   // Find entry in the NULL-terminated string buffer
   LPCWSTR pwstr = m_pwstrNameBuffer;
   if( pwstr == NULL ) return FALSE;
   while( iIndex > 0 ) {
      pwstr += ::lstrlenW(pwstr) + 1;
      iIndex--;
   }
   // Copy name
   USES_CONVERSION;
   ::lstrcpyn(pstrName, W2CT(pwstr), cchMax);
   return TRUE;
}

short COledbRecordset::GetColumnIndex(LPCTSTR pstrName) const
{
   _ASSERTE(IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrName,(UINT)-1));
   USES_CONVERSION;
   LPCWSTR pwstrName = T2CW(pstrName);
   LPCWSTR pwstr = m_pwstrNameBuffer;
   if( pwstr == NULL ) return FALSE;
   for( short i = 0; i < m_nCols; i++ ) {
      if( ::lstrcmpiW(pwstrName, pwstr) == 0 ) return i - m_iAdjustIndex;;
      pwstr += ::lstrlenW(pwstr) + 1;
   }
   return -1;
}

BOOL COledbRecordset::IsEOF() const
{
   if( !IsOpen() ) return TRUE;
   return m_fEOF;
}

BOOL COledbRecordset::MoveNext()
{
   return MoveCursor(0, 1);
}

BOOL COledbRecordset::MovePrev()
{
   return MoveCursor(-2, 1);
}

BOOL COledbRecordset::MoveTop()
{
   _ASSERTE(IsOpen());
   m_spRowset->RestartPosition(DB_NULL_HCHAPTER);
   return MoveCursor(0, 1);
}

BOOL COledbRecordset::MoveBottom()
{
   _ASSERTE(IsOpen());
   // Restart and then move backwards...
   m_spRowset->RestartPosition(DB_NULL_HCHAPTER);
   return MoveCursor(-1, 1);
}

BOOL COledbRecordset::MoveAbs(DWORD dwPos)
{
   _ASSERTE(IsOpen());
   m_spRowset->RestartPosition(DB_NULL_HCHAPTER);
   return MoveCursor(dwPos, 1);
}

DWORD COledbRecordset::GetRowNumber()
{
   _ASSERTE(IsOpen());
   // BUG: Not working because we do not require IRowsetScroll support
   CComQIPtr<IRowsetScroll> spScroll = m_spRowset;
   if( spScroll == NULL ) return 0;
   DBCOUNTITEM nRows;
   BYTE cBookmark = DBBMK_LAST;
   HRESULT Hr = spScroll->GetApproximatePosition(DB_NULL_HCHAPTER, NULL, &cBookmark, NULL, &nRows);
   if( FAILED(Hr) ) return 0;
   return (DWORD)nRows;
}

BOOL COledbRecordset::NextResultset()
{
   _ASSERTE(IsOpen());
   return FALSE;
}

BOOL COledbRecordset::MoveCursor(LONG lSkip, LONG lAmount)
{
   _ASSERTE(IsOpen());
   HRESULT Hr;

   ::ZeroMemory(m_pData, m_dwBufferSize);

   // Move the cursor and get some data
   HROW* rghRows = NULL;
   ULONG nRecevied;
   Hr = m_spRowset->GetNextRows(DB_NULL_HCHAPTER, lSkip, lAmount, &nRecevied, &rghRows);
   // Update EOF marker (HRESULT can be anything but S_OK)
   m_fEOF = Hr != S_OK;
   if( Hr != S_OK ) return TRUE; // Error or reached bottom

   // Get field values
   Hr = m_spRowset->GetData(*rghRows, m_hAccessor, m_pData);

   // Before we check the result, release the rows
   m_spRowset->ReleaseRows(nRecevied, rghRows, NULL, NULL, NULL);

   // Finally, check GetData() result...
   if( FAILED(Hr) ) return _Error(Hr);

   return TRUE;
}

BOOL COledbRecordset::Attach(IRowset* pRowset)
{
   Close();
   m_spRowset = pRowset;
   if( !_BindColumns() ) return TRUE; // In case of UPDATE/INSERT without a result
                                      // then there are no columns
   return MoveNext();
}

BOOL COledbRecordset::_Error(HRESULT Hr)
{
   if( m_pDb == NULL ) return FALSE;
   return m_pDb->_Error(Hr);
}

BOOL COledbRecordset::_BindColumns()
{
   _ASSERTE(m_rgBindings==NULL);
   
   if( !IsOpen() ) return FALSE;

   HRESULT Hr;
   m_nCols = 0;
   m_pwstrNameBuffer = NULL;

   CComQIPtr<IColumnsInfo> spColInfo = m_spRowset;
   if( spColInfo == NULL ) return FALSE;
   DBCOLUMNINFO* rgColumnInfo = NULL;
   ULONG nCols = 0;
   Hr = spColInfo->GetColumnInfo(&nCols, &rgColumnInfo, &m_pwstrNameBuffer);
   if( FAILED(Hr) ) return _Error(Hr);

   // Allocate memory for the bindings array; there is a one-to-one
   // mapping between the columns returned from GetColumnInfo() and our
   // bindings.
   long cbAlloc = nCols * sizeof(DBBINDING);
   m_rgBindings = (DBBINDING*) ::CoTaskMemAlloc(cbAlloc);
   if( m_rgBindings == NULL ) return FALSE;
   ::ZeroMemory(m_rgBindings, cbAlloc);
   m_iAdjustIndex = 0;

   // Construct the binding array element for each column.
   ULONG dwOffset = 0;
   for( ULONG iCol = 0; iCol < nCols; iCol++ ) {
      DBBINDING& b = m_rgBindings[iCol];
      b.iOrdinal = rgColumnInfo[iCol].iOrdinal;
      b.dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;
      b.obStatus = dwOffset;
      b.obLength = dwOffset + sizeof(DBSTATUS);
      b.obValue = dwOffset + sizeof(DBSTATUS) + sizeof(ULONG);    
      b.dwMemOwner = DBMEMOWNER_CLIENTOWNED;
      b.eParamIO = DBPARAMIO_NOTPARAM;
      b.bPrecision = rgColumnInfo[iCol].bPrecision;
      b.bScale = rgColumnInfo[iCol].bScale;

      // Ignore bookmark column
      if( (rgColumnInfo[iCol].dwFlags & DBCOLUMNFLAGS_ISBOOKMARK) != 0 ) m_iAdjustIndex++;

      WORD wType = rgColumnInfo[iCol].wType;
      switch( wType ) {
      case DBTYPE_CY:
      case DBTYPE_DECIMAL:
      case DBTYPE_NUMERIC:
         b.wType = DBTYPE_STR;
         b.cbMaxLen = 50; // Allow 50 characters for conversion
         break;
      case DBTYPE_STR:
      case DBTYPE_WSTR:
#ifdef _UNICODE
         b.wType = DBTYPE_WSTR;
#else
         b.wType = DBTYPE_STR;
#endif
         b.cbMaxLen = max(min((rgColumnInfo[iCol].ulColumnSize + 1UL) * sizeof(TCHAR), 1024UL), 0UL);
         break;
      default:
         b.wType = wType;
         b.cbMaxLen = max(min(rgColumnInfo[iCol].ulColumnSize, 1024UL), 0UL);
      }

// ROUNDUP on all platforms pointers must be aligned properly
#define ROUNDUP_AMOUNT 8
#define ROUNDUP_(size,amount) (((ULONG)(size)+((amount)-1))&~((amount)-1))
#define ROUNDUP(size)         ROUNDUP_(size, ROUNDUP_AMOUNT)
    
      // Update the offset past the end of this column's data
      dwOffset = b.cbMaxLen + b.obValue;
      dwOffset = ROUNDUP(dwOffset);
   }

   m_nCols = (short) nCols;
   m_dwBufferSize = dwOffset;

   ::CoTaskMemFree(rgColumnInfo);

   // Create accessor
   CComQIPtr<IAccessor> spAccessor = m_spRowset;
   if( spAccessor == NULL ) return FALSE;
   Hr = spAccessor->CreateAccessor(DBACCESSOR_ROWDATA, m_nCols, m_rgBindings, 0, &m_hAccessor, NULL);
   if( FAILED(Hr) ) return _Error(Hr);

   m_pData = ::CoTaskMemAlloc(m_dwBufferSize);
   if( m_pData == NULL ) return FALSE;

   return TRUE;
}


//////////////////////////////////////////////////////////////
// COledbCommand
//

COledbCommand::COledbCommand(COledbDatabase* pDb) : 
   m_pDb(pDb), 
   m_rgBindings(NULL), 
   m_hAccessor(DB_NULL_HACCESSOR), 
   m_nParams(0), 
   m_pData(NULL)
{
   _ASSERTE(m_pDb);
   _ASSERTE(m_pDb->IsOpen());
#ifdef _DEBUG
   m_pDb->m_nRecordsets++;
#endif
}

COledbCommand::~COledbCommand()
{
   Close();
#ifdef _DEBUG
   m_pDb->m_nRecordsets--;
#endif
}

BOOL COledbCommand::Create(LPCTSTR pstrSQL, long lType /*= DB_OPEN_TYPE_FORWARD_ONLY*/, long lOptions /*= DB_OPTION_DEFAULT*/)
{
   _ASSERTE(m_pDb->IsOpen());
   _ASSERTE(!::IsBadStringPtr(pstrSQL,(UINT)-1));
   HRESULT Hr;

   Close();

   // Open the command
   CComQIPtr<IDBCreateCommand> spCreate = m_pDb->m_spSession;
   if( spCreate == NULL ) return FALSE;
   CComPtr<ICommand> spCommand;
   Hr = spCreate->CreateCommand(NULL, IID_ICommand, (LPUNKNOWN*) &spCommand);
   if( FAILED(Hr) ) return _Error(Hr);
   COledbDatabase::_SetRecordsetType(spCommand, lType, lOptions);
   m_spText = spCommand;
   _ASSERTE(m_spText);
   USES_CONVERSION;
   Hr = m_spText->SetCommandText(DBGUID_DBSQL, T2COLE(pstrSQL));
   if( FAILED(Hr) ) return _Error(Hr);

   // Prepare the command
   if( (lOptions & DB_OPTION_PREPARE) != 0 ) {
      CComQIPtr<ICommandPrepare> spPrepare = m_spText;
      _ASSERTE(spPrepare);
      if( spPrepare ) {
         Hr = spPrepare->Prepare(0);
         if( FAILED(Hr) ) return _Error(Hr);
      }
   }

   // Create parameter bindings memory area
   long cbAlloc = MAX_PARAMS * sizeof(DBBINDING);
   m_rgBindings = (DBBINDING*) ::CoTaskMemAlloc(cbAlloc);
   if( m_rgBindings == NULL ) return FALSE;
   ::ZeroMemory(m_rgBindings, cbAlloc);

   m_pData = ::CoTaskMemAlloc(MAX_PARAMBUFFER_SIZE);
   _ASSERTE(m_pData);
   if( m_pData == NULL ) return FALSE;
   ::ZeroMemory(m_pData, MAX_PARAMBUFFER_SIZE);

   m_nParams = 0;
   m_dwBindOffset = 0L;

   return TRUE;
}

void COledbCommand::Close()
{
   if( m_rgBindings ) {
      ::CoTaskMemFree(m_rgBindings);
      m_rgBindings = NULL;
   }
   if( m_pData ) {
      ::CoTaskMemFree(m_pData);
      m_pData = NULL;
   }
   m_hAccessor = DB_NULL_HACCESSOR;
   m_spText.Release();
   m_spRowset.Release();
   m_nParams = 0;
}

BOOL COledbCommand::Cancel()
{
   _ASSERTE(IsOpen());
   CComQIPtr<ICommand> spCmd = m_spText;
   if( spCmd == NULL ) return FALSE;
   if( FAILED( spCmd->Cancel() ) ) return FALSE;
   return TRUE;
}

BOOL COledbCommand::Execute(IDbRecordset* pRecordset /*= NULL*/)
{
   _ASSERTE(m_spText);
   HRESULT Hr;

   //CComQIPtr<ICommandWithParameters> spParams = m_spRowset;
   //_ASSERTE(spParams);
   //if( spParams==NULL ) return FALSE;
   //spParams->SetParameterInfo(m_nParams, ...);

   m_nRowsAffected = 0;

   // Create accessor
   CComQIPtr<IAccessor> spAccessor = m_spText;
   if( spAccessor == NULL ) return FALSE;

   if( m_nParams > 0 ) {
#ifdef _DEBUG
      DBBINDSTATUS stat[MAX_PARAMS] = { 0 };
      Hr = spAccessor->CreateAccessor(DBACCESSOR_PARAMETERDATA, m_nParams, m_rgBindings, m_dwBindOffset, &m_hAccessor, stat);
#else
      Hr = spAccessor->CreateAccessor(DBACCESSOR_PARAMETERDATA, m_nParams, m_rgBindings, m_dwBindOffset, &m_hAccessor, NULL);
#endif
      if( FAILED(Hr) ) return _Error(Hr);
   }

   DBPARAMS Params;
   Params.pData = m_pData;
   Params.cParamSets = m_nParams > 0 ? 1 : 0;
   Params.hAccessor = m_hAccessor;

   Hr = m_spText->Execute(NULL, IID_IRowset, &Params, &m_nRowsAffected, (LPUNKNOWN*) &m_spRowset);
   if( FAILED(Hr) ) return _Error(Hr);

   // Did we want to see the result set?
   if( m_spRowset != NULL && pRecordset != NULL ) {
      COledbRecordset* pRec = static_cast<COledbRecordset*>(pRecordset);
      return pRec->Attach(m_spRowset);
   }
   return TRUE;
}
   
BOOL COledbCommand::IsOpen() const
{
   return m_spRowset != NULL;
}

DWORD COledbCommand::GetRowCount() const
{
   return m_nRowsAffected;
}

BOOL COledbCommand::SetParam(short iIndex, const long* pData)
{
   _ASSERTE(m_rgBindings);
   _ASSERTE(iIndex<MAX_PARAMS);
   
   // Fill out binding info
   DBBINDING& b = m_rgBindings[m_nParams];
   b.iOrdinal = iIndex + 1;
   b.wType = DBTYPE_I4;
   b.cbMaxLen = sizeof(long);
   b.dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;
   b.obStatus = m_dwBindOffset;
   b.obLength = m_dwBindOffset + sizeof(DBSTATUS);
   b.obValue = m_dwBindOffset + sizeof(DBSTATUS) + sizeof(ULONG);    
   b.dwMemOwner = DBMEMOWNER_CLIENTOWNED;
   b.eParamIO = DBPARAMIO_INPUT;
   b.bPrecision = 0;
   b.bScale = 0;

   // Set value
   * (long*) ( (LPBYTE) m_pData + b.obValue ) = *pData;

   // Update the offset past the end of this column's data
   m_dwBindOffset = b.cbMaxLen + b.obValue;
   m_dwBindOffset = ROUNDUP(m_dwBindOffset);
   _ASSERTE(m_dwBindOffset<MAX_PARAMBUFFER_SIZE);

   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short) (iIndex + 1);
   return TRUE;
}

BOOL COledbCommand::SetParam(short iIndex, LPCTSTR pData, UINT cchMax /*= -1*/)
{
   _ASSERTE(m_rgBindings);
   _ASSERTE(iIndex<MAX_PARAMS);
   
   int nChars = cchMax == -1 ? ::lstrlen(pData) + 1 : cchMax;
   int nBytes = nChars * sizeof(TCHAR);

   // Fill out binding info
   DBBINDING& b = m_rgBindings[m_nParams];
   b.iOrdinal = iIndex + 1;
#ifdef _UNICODE
   b.wType = DBTYPE_WSTR;
#else
   b.wType = DBTYPE_STR;
#endif
   b.cbMaxLen = nBytes;
   b.dwPart = DBPART_VALUE;
   b.obValue = m_dwBindOffset;
   b.dwMemOwner = DBMEMOWNER_CLIENTOWNED;
   b.eParamIO = DBPARAMIO_INPUT;
   b.bPrecision = 0;
   b.bScale = 0;

   // Set value
   LPTSTR pstr = (LPTSTR) ( (LPBYTE) m_pData + b.obValue );
   ::lstrcpy( pstr, pData );

   // Update the offset past the end of this column's data
   m_dwBindOffset = b.cbMaxLen + b.obValue;
   m_dwBindOffset = ROUNDUP(m_dwBindOffset);
   _ASSERTE(m_dwBindOffset<MAX_PARAMBUFFER_SIZE);

   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short) (iIndex + 1);
   return TRUE;
}

BOOL COledbCommand::SetParam(short /*iIndex*/, const bool* /*pData*/)
{
   _ASSERTE(m_rgBindings);
   // TODO: Code here...
   _ASSERTE(false);
   return FALSE;
}

BOOL COledbCommand::SetParam(short iIndex, const float* pData)
{
   _ASSERTE(m_rgBindings);
   _ASSERTE(iIndex<MAX_PARAMS);
   
   // Fill out binding info
   DBBINDING& b = m_rgBindings[m_nParams];
   b.iOrdinal = iIndex + 1;
   b.wType = DBTYPE_R4;
   b.cbMaxLen = sizeof(float);
   b.dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;
   b.obStatus = m_dwBindOffset;
   b.obLength = m_dwBindOffset + sizeof(DBSTATUS);
   b.obValue = m_dwBindOffset + sizeof(DBSTATUS) + sizeof(ULONG);    
   b.dwMemOwner = DBMEMOWNER_CLIENTOWNED;
   b.eParamIO = DBPARAMIO_INPUT;
   b.bPrecision = 0;
   b.bScale = 0;

   // Set value
   * (float*) ( (LPBYTE)m_pData + b.obValue ) = *pData;

   // Update the offset past the end of this column's data
   m_dwBindOffset = b.cbMaxLen + b.obValue;
   m_dwBindOffset = ROUNDUP(m_dwBindOffset);
   _ASSERTE(m_dwBindOffset<MAX_PARAMBUFFER_SIZE);

   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short) (iIndex + 1);
   return TRUE;
}

BOOL COledbCommand::SetParam(short iIndex, const double* pData)
{
   _ASSERTE(m_rgBindings);
   _ASSERTE(iIndex<MAX_PARAMS);
   
   // Fill out binding info
   DBBINDING& b = m_rgBindings[m_nParams];
   b.iOrdinal = iIndex + 1;
   b.wType = DBTYPE_R8;
   b.cbMaxLen = sizeof(double);
   b.dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;
   b.obStatus = m_dwBindOffset;
   b.obLength = m_dwBindOffset + sizeof(DBSTATUS);
   b.obValue = m_dwBindOffset + sizeof(DBSTATUS) + sizeof(ULONG);    
   b.dwMemOwner = DBMEMOWNER_CLIENTOWNED;
   b.eParamIO = DBPARAMIO_INPUT;
   b.bPrecision = 0;
   b.bScale = 0;

   // Set value
   * (double*) ( (LPBYTE)m_pData + b.obValue ) = *pData;

   // Update the offset past the end of this column's data
   m_dwBindOffset = b.cbMaxLen + b.obValue;
   m_dwBindOffset = ROUNDUP(m_dwBindOffset);
   _ASSERTE(m_dwBindOffset<MAX_PARAMBUFFER_SIZE);

   // Manage parameter count
   if( iIndex >= m_nParams ) m_nParams = (short) (iIndex + 1);
   return TRUE;
}

BOOL COledbCommand::SetParam(short /*iIndex*/, const SYSTEMTIME* /*pData*/, short /*iType*/)
{
   _ASSERTE(m_rgBindings);
   // TODO: Code here...
   _ASSERTE(false);
   return FALSE;
}

#if defined(_STRING_)

BOOL COledbCommand::SetParam(short iIndex, std::string& str)
{
   return SetParam(iIndex, str.c_str());
}

#endif // _STRING

#if defined(__ATLSTR_H__) || defined(_WTL_USE_CSTRING)

BOOL COledbCommand::SetParam(short iIndex, CString& str)
{
   return SetParam(iIndex, (LPCTSTR) str);
}

#endif // __ATLSTR_H__

BOOL COledbCommand::_Error(HRESULT Hr)
{
   _ASSERTE(m_pDb);
   return m_pDb->_Error(Hr);
}


//////////////////////////////////////////////////////////////
// COledbErrors
//

COledbErrors::COledbErrors() 
   : m_lCount(0L)
{
}

COledbErrors::~COledbErrors()
{
}

void COledbErrors::_Init(HRESULT Hr)
{
   m_lCount = 0;

   CComPtr<IErrorInfo> spErrorInfo;
   ::GetErrorInfo(0, &spErrorInfo);
   if( spErrorInfo == NULL ) {
      // No error object
      m_p[0]._Init(Hr, CComBSTR(L"System"), CComBSTR(L"Function failed."));
      m_lCount = 1;
      return;
   }

   CComQIPtr<IErrorRecords> spErrorRecs = spErrorInfo;
   if( spErrorRecs != NULL ) {
      // Provider supports IErrorRecords
      ULONG nCount = 0;
      spErrorRecs->GetRecordCount(&nCount);
      for( ULONG i = 0; i < nCount; i++ ) {
         HRESULT Hr;
         ERRORINFO ErrorInfo;
         Hr = spErrorRecs->GetBasicErrorInfo(i, &ErrorInfo);
         if( FAILED(Hr) ) break;
         CComPtr<IErrorInfo> spErrorInfoRec;
         Hr = spErrorRecs->GetErrorInfo(i, ::GetUserDefaultLCID(), &spErrorInfoRec);
         if( FAILED(Hr) ) break;

         CComBSTR bstrSource;
         spErrorInfoRec->GetSource(&bstrSource);
         CComBSTR bstrMsg;
         spErrorInfoRec->GetDescription(&bstrMsg);
      
         m_p[m_lCount]._Init(ErrorInfo.hrError, bstrSource, bstrMsg);
         m_lCount++;
         if( m_lCount >= MAX_ERRORS ) break;
      }
   }

   if( m_lCount == 0 ) {
      // Provider only supports IErrorInfo
      CComBSTR bstrSource;
      spErrorInfo->GetSource(&bstrSource);
      CComBSTR bstrMsg;
      spErrorInfo->GetDescription(&bstrMsg);
      m_p[m_lCount]._Init(Hr, bstrSource, bstrMsg);
      m_lCount++;
   }
}

long COledbErrors::GetCount() const
{
   return m_lCount;
}

void COledbErrors::Clear()
{
   m_lCount = 0;
}

IDbError* COledbErrors::GetError(short iIndex)
{
   _ASSERTE(iIndex>=0 && iIndex<m_lCount);
   if( iIndex < 0 || iIndex >= m_lCount ) return NULL;
   return &m_p[iIndex];
}


//////////////////////////////////////////////////////////////
// COledbError
//

COledbError::COledbError()
{
}

COledbError::~COledbError()
{
}

void COledbError::_Init(long lNativeCode, BSTR bstrSource, BSTR bstrMsg)
{
   m_lNative = lNativeCode;
   m_bstrSource = bstrSource;
   m_bstrMsg = bstrMsg;
}

long COledbError::GetErrorCode()
{
   return m_lNative;
}

long COledbError::GetNativeErrorCode()
{
   return m_lNative;
}

void COledbError::GetOrigin(LPTSTR pstrStr, UINT cchMax)
{
   USES_CONVERSION;
   ::lstrcpyn(pstrStr, OLE2CT(m_bstrSource), cchMax);
}

void COledbError::GetMessage(LPTSTR pstrStr, UINT cchMax)
{
   USES_CONVERSION;
   ::lstrcpyn(pstrStr, OLE2CT(m_bstrMsg), cchMax);
}

void COledbError::GetSource(LPTSTR pstrStr, UINT cchMax)
{
   USES_CONVERSION;
   ::lstrcpyn(pstrStr, OLE2CT(m_bstrSource), cchMax);
}
