#pragma once

#ifndef STRICT
#define STRICT
#endif


#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit
#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW


#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
#include <atlstr.h>
#include <ShlObj.h>



void DbgReport(_In_opt_z_ char const* format,...);
ATL::CStringW DumpIID(_In_ REFIID iid) throw();
HRESULT _ShellFolderFromPidl(PCUIDLIST_RELATIVE pidl, ATL::CComPtr<IShellFolder>& folder);


#define DBG_TRACE(fmt, ...) DbgReport(fmt, __VA_ARGS__)

