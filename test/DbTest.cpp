// DbTest.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

// The ODBC / OLEDB test requires a pre-created "Northwind" datasource.
// Usually you would just create an ODBC MS Access entry for the Northwind.MDB
// database that comes with many Microsoft tools (MS Access & VB6 to name a few).
#include "DbOdbc.h"
#include "DbOledb.h"

// The CSV test loads a sample file located in the source folder
#include "DbCsv.h"

// To remove SQL Lite from the test, comment the following line out
// and remove the cSqlite.cpp from the project...
//#include "DbSqlite2.h"
//#include "DbSqlite3.h"


void OdbcTest(LPCTSTR pstrConnection)
{
   COdbcSystem System;
   System.Initialize();

   COdbcDatabase Db(&System);
   BOOL bRes = Db.Open(NULL, pstrConnection, _T(""), _T(""));
   if( !bRes ) return;
   COdbcRecordset Rec(&Db);
   bRes = Rec.Open(_T("SELECT ProductID,ProductName FROM Products"));
   if( !bRes ) return;
   while( !Rec.IsEOF() ) {
      long lID;
      TCHAR szName[128];
      Rec.GetField(0, lID);
      Rec.GetField(1, szName, 128);
      Rec.MoveNext();
   }
   Rec.Close();
   Db.Close();

   System.Terminate();
}


void Test1(IDbSystem* pSystem, LPCTSTR pstrConnection)
{
   pSystem->Initialize();

   CAutoPtr<IDbDatabase> pDb(pSystem->CreateDatabase());
   BOOL bRes;
   bRes = pDb->Open(NULL, pstrConnection, _T(""), _T(""), DB_OPEN_READ_ONLY);
   if( !bRes ) {
      TCHAR szMsg[256];
      pDb->GetErrors()->GetError(0)->GetMessage(szMsg, 256);
      ::MessageBox( NULL, szMsg, _T("Database Test"), MB_OK|MB_ICONERROR);
      return;
   }

   CAutoPtr<IDbRecordset> pRec(pSystem->CreateRecordset(pDb));
   bRes = pRec->Open(_T("ProductID,ProductName,UnitPrice,Discontinued FROM Products ORDER BY ProductID"), DB_OPEN_TYPE_DYNASET);
   
   DWORD dwCnt = pRec->GetRowCount();
   dwCnt;

   while( !pRec->IsEOF() ) {
      long lID;
      TCHAR szTitle[128];
      float fUnitPrice;
      double dblUnitPrice;
      bool bDiscontinued;
      pRec->GetField(0, lID);
      pRec->GetField(1, szTitle, 128);
      pRec->GetField(2, fUnitPrice);
      pRec->GetField(2, dblUnitPrice);
      pRec->GetField(3, bDiscontinued);
      pRec->MoveNext();
      //DWORD dwRow = pRec->GetRowNumber();
   }

   DWORD nFields = pRec->GetColumnCount();
   nFields;
   long iIndex = pRec->GetColumnIndex(_T("UnitPrice"));
   iIndex;

   pRec->Close();

   pDb->Close();

   pSystem->Terminate();
}

void Test2(IDbSystem* pSystem, LPCTSTR pstrConnection)
{
   pSystem->Initialize();

   CAutoPtr<IDbDatabase> pDb(pSystem->CreateDatabase());
   BOOL bRes;
   bRes = pDb->Open(NULL, pstrConnection, _T(""), _T(""), DB_OPEN_READ_ONLY);
   if( !bRes ) {
      TCHAR szMsg[256];
      pDb->GetErrors()->GetError(0)->GetMessage(szMsg, 256);
      ::MessageBox( NULL, szMsg, _T("Database Test"), MB_OK|MB_ICONERROR);
      return;
   }

   CAutoPtr<IDbCommand> pCmd(pSystem->CreateCommand(pDb));
   pCmd->Create(_T("SELECT ProductID,ProductName FROM Products WHERE ProductID=? AND ProductName=?"));
   long lVal = 2L;
   pCmd->SetParam(0, &lVal);
   pCmd->SetParam(1, _T("Chang"));
   CAutoPtr<IDbRecordset> pRec(pSystem->CreateRecordset(pDb));
   pCmd->Execute(pRec);
   while( !pRec->IsEOF() ) {
      long lID;
      TCHAR szTitle[128];
      pRec->GetField(0, lID);
      pRec->GetField(1, szTitle, 128);
      pRec->MoveNext();
   }
   pRec->Close();
   pCmd->Close();

   pDb->Close();

   pSystem->Terminate();
}


void CsvTest(LPCTSTR pstrFilename)
{
   TCHAR szFilename[MAX_PATH];
   ::GetModuleFileName(NULL, szFilename, MAX_PATH);
   LPTSTR p = _tcsrchr(szFilename, '\\');
   _tcscpy(p + 1, pstrFilename);
   CCsvSystem system;
   CCsvDatabase db(&system);
   db.Open(NULL, szFilename, NULL, NULL);
   CCsvRecordset rec(&db);
   rec.Open(_T(""));
   while( !rec.IsEOF() ) {
      TCHAR szValue[128];
      long lValue;
      rec.GetField(0, szValue, 128);
      rec.GetField(7, lValue);
      rec.MoveNext();
   }
   rec.Close();
   db.Close();
}


int main(int /*argc*/, char* /*argv[]*/)
{
    LPCTSTR pstrCSV = _T("../../data/citylist.csv");

    printf("Testing csv...");
    CsvTest(pstrCSV);
    printf("Done.\n");

    LPCTSTR pstrDSN = _T("Northwind");

    printf("Testing ODBC...");
    OdbcTest(pstrDSN);
    printf("Done.\n");

    pstrDSN = _T("Provider=SQLOLEDB;Data Source=(local);Integrated Security=SSPI;Initial Catalog=Northwind");

    printf("Testing ODBC...");
    CAutoPtr<IDbSystem> pODBC(new COdbcSystem());
    Test1(pODBC, pstrDSN);
    Test2(pODBC, pstrDSN);
    printf("Done.\n");

    printf("Testing OLEDB...");
    CAutoPtr<IDbSystem> pOLEDB(new COledbSystem());
    Test1(pOLEDB, pstrDSN);
    Test2(pOLEDB, pstrDSN);
    printf("Done.\n");

    return 0;
}
