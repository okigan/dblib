#include "stdafx.h"
#include "dblib/dblib.h"
#include "dbcsv.h"
#include "dbodbc.h"
#include "dboledb.h"

#ifdef USE_SQLITE_2
# include "dbsqlite2.h"
#endif

#ifdef USE_SQLITE_3
# include "dbsqlite3.h"
#endif

CComModule _Module;

BOOL OpenDbSystem(long reserved, ENUM_DB_SYSTEM eSystem, IDbSystem** ppiDbSystem)
{
  IDbSystem* pi = NULL;
  switch(eSystem)
  {
    //TODO: move to enum or etc
  case DB_SYSTEM_CVS:
    {
      CComObject<CCsvSystem>* obj = NULL;
      CComObject<CCsvSystem>::CreateInstance(&obj);
      obj->Initialize();
      pi = obj;
    }break;
  case DB_SYSTEM_ODBC:
    {
      CComObject<COdbcSystem>* obj = NULL;
      CComObject<COdbcSystem>::CreateInstance(&obj);
      obj->Initialize();
      pi = obj;
    }
    break;
  case DB_SYSTEM_OLEDB:
    {
      CComObject<COledbSystem>* obj = NULL;
      CComObject<COledbSystem>::CreateInstance(&obj);
      obj->Initialize();
      pi = obj;
    }
    break;
  }

  if(*ppiDbSystem = pi){
    return TRUE;
  }else{
    return FALSE;
  }
}