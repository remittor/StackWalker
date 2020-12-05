/**********************************************************************
 *
 * StackWalker.cpp
 * https://github.com/JochenKalmbach/StackWalker
 *
 * Old location: http://stackwalker.codeplex.com/
 *
 *
 * History:
 *  2005-07-27   v1    - First public release on http://www.codeproject.com/
 *                       http://www.codeproject.com/threads/StackWalker.asp
 *  2005-07-28   v2    - Changed the params of the constructor and ShowCallstack
 *                       (to simplify the usage)
 *  2005-08-01   v3    - Changed to use 'CONTEXT_FULL' instead of CONTEXT_ALL
 *                       (should also be enough)
 *                     - Changed to compile correctly with the PSDK of VC7.0
 *                       (GetFileVersionInfoSizeA and GetFileVersionInfoA is wrongly defined:
 *                        it uses LPSTR instead of LPCSTR as first parameter)
 *                     - Added declarations to support VC5/6 without using 'dbghelp.h'
 *                     - Added a 'pUserData' member to the ShowCallstack function and the
 *                       PReadProcessMemoryRoutine declaration (to pass some user-defined data,
 *                       which can be used in the readMemoryFunction-callback)
 *  2005-08-02   v4    - OnSymInit now also outputs the OS-Version by default
 *                     - Added example for doing an exception-callstack-walking in main.cpp
 *                       (thanks to owillebo: http://www.codeproject.com/script/profile/whos_who.asp?id=536268)
 *  2005-08-05   v5    - Removed most Lint (http://www.gimpel.com/) errors... thanks to Okko Willeboordse!
 *  2008-08-04   v6    - Fixed Bug: Missing LEAK-end-tag
 *                       http://www.codeproject.com/KB/applications/leakfinder.aspx?msg=2502890#xx2502890xx
 *                       Fixed Bug: Compiled with "WIN32_LEAN_AND_MEAN"
 *                       http://www.codeproject.com/KB/applications/leakfinder.aspx?msg=1824718#xx1824718xx
 *                       Fixed Bug: Compiling with "/Wall"
 *                       http://www.codeproject.com/KB/threads/StackWalker.aspx?msg=2638243#xx2638243xx
 *                       Fixed Bug: Now checking SymUseSymSrv
 *                       http://www.codeproject.com/KB/threads/StackWalker.aspx?msg=1388979#xx1388979xx
 *                       Fixed Bug: Support for recursive function calls
 *                       http://www.codeproject.com/KB/threads/StackWalker.aspx?msg=1434538#xx1434538xx
 *                       Fixed Bug: Missing FreeLibrary call in "GetModuleListTH32"
 *                       http://www.codeproject.com/KB/threads/StackWalker.aspx?msg=1326923#xx1326923xx
 *                       Fixed Bug: SymDia is number 7, not 9!
 *  2008-09-11   v7      For some (undocumented) reason, dbhelp.h is needing a packing of 8!
 *                       Thanks to Teajay which reported the bug...
 *                       http://www.codeproject.com/KB/applications/leakfinder.aspx?msg=2718933#xx2718933xx
 *  2008-11-27   v8      Debugging Tools for Windows are now stored in a different directory
 *                       Thanks to Luiz Salamon which reported this "bug"...
 *                       http://www.codeproject.com/KB/threads/StackWalker.aspx?msg=2822736#xx2822736xx
 *  2009-04-10   v9      License slightly corrected (<ORGANIZATION> replaced)
 *  2009-11-01   v10     Moved to http://stackwalker.codeplex.com/
 *  2009-11-02   v11     Now try to use IMAGEHLP_MODULE64_V3 if available
 *  2010-04-15   v12     Added support for VS2010 RTM
 *  2010-05-25   v13     Now using secure MyStrcCpy. Thanks to luke.simon:
 *                       http://www.codeproject.com/KB/applications/leakfinder.aspx?msg=3477467#xx3477467xx
 *  2013-01-07   v14     Runtime Check Error VS2010 Debug Builds fixed:
 *                       http://stackwalker.codeplex.com/workitem/10511
 *
 *
 * LICENSE (http://www.opensource.org/licenses/bsd-license.php)
 *
 *   Copyright (c) 2005-2013, Jochen Kalmbach
 *   All rights reserved.
 *
 *   Redistribution and use in source and binary forms, with or without modification,
 *   are permitted provided that the following conditions are met:
 *
 *   Redistributions of source code must retain the above copyright notice,
 *   this list of conditions and the following disclaimer.
 *   Redistributions in binary form must reproduce the above copyright notice,
 *   this list of conditions and the following disclaimer in the documentation
 *   and/or other materials provided with the distribution.
 *   Neither the name of Jochen Kalmbach nor the names of its contributors may be
 *   used to endorse or promote products derived from this software without
 *   specific prior written permission.
 *   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 *   AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
 *   THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 *   ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE
 *   FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 *   (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 *   LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 *   ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 *   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 *   SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 **********************************************************************/

#include "StackWalker.h"

#include <stdio.h>
#include <stdlib.h>
#include <windows.h>
#include <errno.h>
#include <stdarg.h>
#include <new>

#pragma comment(lib, "version.lib") // for "VerQueryValue"

// Normally it should be enough to use 'CONTEXT_FULL' (better would be 'CONTEXT_ALL')
#define STKWLK_CONTEXT_FLAGS CONTEXT_FULL

#pragma warning(disable : 4826)
#if _MSC_VER >= 1900
#pragma warning(disable : 4091)   // For fix unnamed enums from DbgHelp.h
#endif

#if !defined(SE_IMPERSONATE_NAME) || !defined(IMAGE_DLLCHARACTERISTICS_NO_ISOLATION)
#error "Required Microsoft Platform SDK v5.2.3718.0 (November 2002) or later"
// PSDK v5.2.3790.0 (Feb 2003) is last version with VC6 support and Windows 9X support.
#endif

#ifdef DBGHELP_TRANSLATE_TCHAR
#undef DBGHELP_TRANSLATE_TCHAR
#endif

#pragma pack(push, 8)
#include <dbghelp.h>
#pragma pack(pop)

#ifdef StackWalk
#undef StackWalk
#endif

#ifdef SearchPath
#undef SearchPath
#endif

#ifdef SymLoadModule
#undef SymLoadModule
#endif

#ifdef SymFromAddr
#undef SymFromAddr
#endif

#ifdef _UNICODE
  typedef SYMBOL_INFOW        T_SYMBOL_INFO;
  typedef IMAGEHLP_LINEW64    T_IMAGEHLP_LINE64;
#else
  typedef SYMBOL_INFO         T_SYMBOL_INFO;
  typedef IMAGEHLP_LINE64     T_IMAGEHLP_LINE64;
#endif

#ifndef _countof
#define _countof(_Array) (sizeof(_Array) / sizeof(_Array[0]))
#endif

// secure-CRT_functions are only available starting with VC8 (MSVC2005)
#if _MSC_VER < 1400
#ifdef _TRUNCATE
#undef _TRUNCATE
#endif
#define _TRUNCATE  ((SIZE_T)-1)
#define errno_t  int
static errno_t strncpy_s(char * dst, size_t dstcap, const char * src, size_t count)
{
  strncpy(dst, src, min(count, dstcap));
  dst[dstcap-1] = 0;
  return 0;
}
static errno_t strcpy_s(char * dst, size_t dstcap, const char * src)
{
  return strncpy_s(dst, dstcap, src, _TRUNCATE);
}
static errno_t strncat_s(char * dst, size_t dstcap, const char * src, size_t count)
{
  size_t dlen = strlen(dst);
  if (dlen + 1 >= dstcap)
    return 0;
  return strncpy_s(dst + dlen, dstcap - dlen - 1, src, count);
}
static errno_t strcat_s(char * dst, size_t dstcap, const char * src)
{
  return strncat_s(dst, dstcap, src, _TRUNCATE);
}
static errno_t wcsncpy_s(WCHAR * dst, size_t dstcap, const WCHAR * src, size_t count)
{
  wcsncpy(dst, src, min(count, dstcap));
  dst[dstcap-1] = 0;
  return 0;
}
static errno_t wcscpy_s(WCHAR * dst, size_t dstcap, const WCHAR * src)
{
  return wcsncpy_s(dst, dstcap, src, _TRUNCATE);
}
static errno_t wcsncat_s(WCHAR * dst, size_t dstcap, const WCHAR * src, size_t count)
{
  size_t dlen = wcslen(dst);
  if (dlen + 1 >= dstcap)
    return 0;
  return wcsncpy_s(dst + dlen, dstcap - dlen - 1, src, count);
}
static errno_t wcscat_s(WCHAR * dst, size_t dstcap, const WCHAR * src)
{
  return wcsncat_s(dst, dstcap, src, _TRUNCATE);
}
#define _snprintf_s  _snprintf
#define _sntprintf_s _sntprintf
#endif

static errno_t MyStrCpy(LPSTR szDest, size_t nMaxDestSize, LPCSTR szSrc) STKWLK_NOEXCEPT
{
  if (nMaxDestSize == 0 || szSrc == NULL)
    return EINVAL;
  errno_t rc = strncpy_s(szDest, nMaxDestSize, szSrc, _TRUNCATE);
  // INFO: _TRUNCATE will ensure that it is null-terminated;
  // but with older compilers (<1400) it uses "strncpy" and this does not!)
  szDest[nMaxDestSize - 1] = 0;
  return rc;
}

static errno_t MyStrCpy(LPWSTR szDest, size_t nMaxDestSize, LPCWSTR szSrc) STKWLK_NOEXCEPT
{
  if (nMaxDestSize == 0 || szSrc == NULL)
    return EINVAL;
  errno_t rc = wcsncpy_s(szDest, nMaxDestSize, szSrc, _TRUNCATE);
  szDest[nMaxDestSize - 1] = 0;
  return rc;
}

static errno_t MyStrCat(LPSTR szDest, size_t nMaxDestSize, LPCSTR szSrc) STKWLK_NOEXCEPT
{
  if (nMaxDestSize == 0 || szSrc == NULL)
    return EINVAL;
  errno_t rc = strncat_s(szDest, nMaxDestSize, szSrc, _TRUNCATE);
  szDest[nMaxDestSize - 1] = 0;
  return rc;
}

static errno_t MyStrCat(LPWSTR szDest, size_t nMaxDestSize, LPCWSTR szSrc) STKWLK_NOEXCEPT
{
  if (nMaxDestSize == 0 || szSrc == NULL)
    return EINVAL;
  errno_t rc = wcsncat_s(szDest, nMaxDestSize, szSrc, _TRUNCATE);
  szDest[nMaxDestSize - 1] = 0;
  return rc;
}

static errno_t MyTStrFmt(LPTSTR dst, size_t dstcap, LPCTSTR fmt, ...) STKWLK_NOEXCEPT
{
  if (dst == NULL || dstcap < 2 || fmt == NULL)
    return EINVAL;
  memset(dst, 0, dstcap * sizeof(TCHAR));
  va_list argptr;
  va_start(argptr, fmt);
#if _MSC_VER >= 1400
  _vsntprintf_s(dst, dstcap, _TRUNCATE, fmt, argptr);
#else
  _vsntprintf(dst, dstcap, fmt, argptr);
#endif
  va_end(argptr);
  dst[dstcap - 1] = 0;
  return 0;
}

static LPVOID GetProcAddrEx(int & counter, HMODULE hLib, LPCSTR name, LPVOID * ptr = NULL) STKWLK_NOEXCEPT
{
  LPVOID proc = GetProcAddress(hLib, name);
  counter += proc ? 1 : 0;
  if (ptr)
    *ptr = proc;
  return proc;
}

static errno_t MyPathCat(LPWSTR path, size_t capacity, LPCWSTR addon, bool isDir) STKWLK_NOEXCEPT
{
  errno_t rc = MyStrCat(path, capacity, addon);
  if (rc)
    return rc;
  size_t len = wcslen(path);
  if (path[len - 1] == L'\\')
    return 0;
  if (isDir)
    return MyStrCat(path, capacity, L"\\");
  LPWSTR p = wcsrchr(path, L'\\');
  if (p)
    p[1] = 0;
  return 0;
}

static HMODULE LoadDbgHelpLib(bool prefixIsDir, LPCWSTR prefix, LPCWSTR path) STKWLK_NOEXCEPT
{
  WCHAR buf[2048];
  wcscpy_s(buf, _countof(buf), L"\\\\?\\");   // for long path support
  if (prefix)
  {
    errno_t rc = MyPathCat(buf, _countof(buf), prefix, prefixIsDir);
    if (rc)
      return NULL;
  }
  if (path)
  {
    if (path[0] == L'\\')
      path++;
    errno_t rc = MyPathCat(buf, _countof(buf), path, true);
    if (rc)
      return NULL;
  }
  errno_t rc = MyStrCat(buf, _countof(buf), L"dbghelp.dll");
  if (rc)
    return NULL;
  size_t len = wcslen(buf);
  LPCWSTR libpath = (len < MAX_PATH) ? buf + 4 : buf;
  if (GetFileAttributesW(libpath) == INVALID_FILE_ATTRIBUTES)
    return NULL;
  HMODULE hLib = LoadLibraryW(libpath);
  if (!hLib)
    return NULL;
  LPVOID fn = GetProcAddress(hLib, "StackWalk64");
  if (fn != NULL)
    return hLib;
  FreeLibrary(hLib);
  return NULL;
}


class StackWalkerInternal
{
public:
  StackWalkerInternal(StackWalkerBase * parent, HANDLE hProcess, PCONTEXT ctx) STKWLK_NOEXCEPT
  {
    m_parent = parent;
    m_hDbhHelp = NULL;
    m_hProcess = hProcess;
    m_SymInitialized = FALSE;
    m_IHM64Version = 0;      // unknown version
    memset(&Sym, 0, sizeof(Sym));
    m_ctx.ContextFlags = 0;
    if (ctx != NULL)
      m_ctx = *ctx;
  }

  ~StackWalkerInternal() STKWLK_NOEXCEPT
  {
    UnloadDbgHelpLib();
    m_parent = NULL;
  }

  void UnloadDbgHelpLib() STKWLK_NOEXCEPT
  {
    if (m_hDbhHelp != NULL && m_SymInitialized != FALSE && Sym.Cleanup != NULL)
      Sym.Cleanup(m_hProcess);
    memset(&Sym, 0, sizeof(Sym));
    m_SymInitialized = FALSE;
    if (m_hDbhHelp == NULL)
      return;
    FreeLibrary(m_hDbhHelp);
    m_hDbhHelp = NULL;
  }

  BOOL Init(LPCTSTR szSymPath) STKWLK_NOEXCEPT
  {
    TCHAR buf[StackWalkerBase::STACKWALK_MAX_NAMELEN];

    if (m_parent == NULL)
      return FALSE;

    if (!m_hDbhHelp && m_parent->m_szDbgHelpPath)
      m_hDbhHelp = LoadLibraryW(m_parent->m_szDbgHelpPath);

    // Dynamically load the Entry-Points for dbghelp.dll:
    // First try to load the newest one from
    WCHAR szTemp[4096];
    // But before we do this, we first check if the ".local" file exists
    size_t len = GetModuleFileNameW(NULL, szTemp, _countof(szTemp));
    if (!m_hDbhHelp && len > 0 && len < _countof(szTemp)-64)
    {
      MyStrCat(szTemp, _countof(szTemp), L".local");
      if (GetFileAttributesW(szTemp) == INVALID_FILE_ATTRIBUTES)
      {
        // ".local" file does not exist, so we can try to load the dbghelp.dll from the "Debugging Tools for Windows"
        // Ok, first try the new path according to the architecture:
        WCHAR szProgDir[MAX_PATH];
        for (int i = 0; i <= 1; i++) {
          LPCWSTR envname = (i == 0) ? L"ProgramFiles" : L"ProgramFiles(x86)";
          size_t pdlen = GetEnvironmentVariableW(envname, szProgDir, _countof(szProgDir));
          if (pdlen == 0 || pdlen >= _countof(szProgDir)-1)
            continue;
#ifdef _M_IX86
          m_hDbhHelp = LoadDbgHelpLib(true, szProgDir, L"Windows Kits\\10\\Debuggers\\x86");
#elif _M_X64
          m_hDbhHelp = LoadDbgHelpLib(true, szProgDir, L"Windows Kits\\10\\Debuggers\\x64");
#elif _M_IA64
          m_hDbhHelp = LoadDbgHelpLib(true, szProgDir, L"Windows Kits\\10\\Debuggers\\ia64");
#endif
          if (m_hDbhHelp)
            break;
#ifdef _M_IX86
          m_hDbhHelp = LoadDbgHelpLib(true, szProgDir, L"Debugging Tools for Windows (x86)");
#elif _M_X64
          m_hDbhHelp = LoadDbgHelpLib(true, szProgDir, L"Debugging Tools for Windows (x64)");
#elif _M_IA64
          m_hDbhHelp = LoadDbgHelpLib(true, szProgDir, L"Debugging Tools for Windows (ia64)");
#endif
          if (m_hDbhHelp)
            break;
          // If still not found, try the old directories...
          m_hDbhHelp = LoadDbgHelpLib(true, szProgDir, L"Debugging Tools for Windows");
          if (m_hDbhHelp)
            break;
#if defined _M_X64 || defined _M_IA64
          // Still not found? Then try to load the (old) 64-Bit version:
          m_hDbhHelp = LoadDbgHelpLib(true, szProgDir, L"Debugging Tools for Windows 64-Bit");
          if (m_hDbhHelp)
            break;
#endif
        } // for
      }
    }
    if (m_hDbhHelp == NULL) // if not already loaded, try to load a default-one
      m_hDbhHelp = LoadLibraryW(L"dbghelp.dll");

    if (m_hDbhHelp == NULL)
      return FALSE;

    ULONGLONG verFile = 0;
    memset(buf, 0, sizeof(buf));
    DWORD dwLen = GetModuleFileName(m_hDbhHelp, buf, _countof(buf)-1);
    buf[(dwLen == 0) ? 0 : _countof(buf)-1] = 0;
    if (_tcslen(buf) > 0)
      GetFileVersion(buf, verFile);
    this->m_parent->OnLoadDbgHelp(verFile, buf);

    memset(&Sym, 0, sizeof(Sym));
    int fcnt = 0;
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymCleanup", (LPVOID*)&Sym.Cleanup);
#ifdef _UNICODE
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymInitializeW", (LPVOID*)&Sym.Initialize);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetModuleInfoW64", (LPVOID*)&Sym.GetModuleInfo);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymFromAddrW", (LPVOID*)&Sym.FromAddr);
    GetProcAddrEx(fcnt, m_hDbhHelp, "UnDecorateSymbolNameW", (LPVOID*)&Sym.UnDecorateName);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymLoadModuleExW", (LPVOID*)&Sym.LoadModuleEx);
#else
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymInitialize", (LPVOID*)&Sym.Initialize);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetModuleInfo64", (LPVOID*)&Sym.GetModuleInfo);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetSymFromAddr64", (LPVOID*)&Sym.GetSymFromAddr);
    GetProcAddrEx(fcnt, m_hDbhHelp, "UnDecorateSymbolName", (LPVOID*)&Sym.UnDecorateName);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymLoadModule64", (LPVOID*)&Sym.LoadModule);
#endif
    GetProcAddrEx(fcnt, m_hDbhHelp, "StackWalk64", (LPVOID*)&Sym.StackWalk);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetOptions", (LPVOID*)&Sym.GetOptions);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymSetOptions", (LPVOID*)&Sym.SetOptions);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymFunctionTableAccess64", (LPVOID*)&Sym.FunctionTableAccess);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetModuleBase64", (LPVOID*)&Sym.GetModuleBase);

    if (fcnt < 11)
    {
      this->m_parent->OnDbgHelpErr(_T("LoadDbgHelp"), ERROR_INVALID_TABLE, 0);
      UnloadDbgHelpLib();
      return FALSE;
    }
    fcnt = 0;
#ifdef _UNICODE
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetLineFromAddrW64", (LPVOID*)&Sym.GetLineFromAddr);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetSearchPathW", (LPVOID*)&Sym.GetSearchPath);
#else
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetLineFromAddr64", (LPVOID*)&Sym.GetLineFromAddr);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymGetSearchPath", (LPVOID*)&Sym.GetSearchPath);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymLoadModuleEx", (LPVOID*)&Sym.LoadModuleEx);
    GetProcAddrEx(fcnt, m_hDbhHelp, "SymFromAddr", (LPVOID*)&Sym.FromAddr);
#endif

    m_SymInitialized = Sym.Initialize(m_hProcess, szSymPath, FALSE);
    if (m_SymInitialized == FALSE)
    {
      this->m_parent->OnDbgHelpErr(_T("SymInitialize"), GetLastError(), 0);
      UnloadDbgHelpLib();
      return FALSE;
    }

    DWORD symOptions = Sym.GetOptions();
    symOptions |= SYMOPT_LOAD_LINES;
    symOptions |= SYMOPT_FAIL_CRITICAL_ERRORS;
    //symOptions |= SYMOPT_NO_PROMPTS;
    symOptions = Sym.SetOptions(symOptions);

    memset(buf, 0, sizeof(buf));
    if (Sym.GetSearchPath != NULL)
    {
      if (Sym.GetSearchPath(m_hProcess, buf, _countof(buf) - 1) == FALSE)
        this->m_parent->OnDbgHelpErr(_T("SymGetSearchPath"), GetLastError(), 0);
    }
    TCHAR szUserName[1024] = {0};
    DWORD dwSize = 1024;
    GetUserName(szUserName, &dwSize);
    this->m_parent->OnSymInit(buf, symOptions, szUserName);

    return TRUE;
  }

  StackWalkerBase * m_parent;

  CONTEXT m_ctx;
  HMODULE m_hDbhHelp;
  HANDLE  m_hProcess;
  BOOL    m_SymInitialized;
  char    m_IHM64Version;      // actual version of IMAGEHLP_MODULE64 struct

#pragma pack(push, 8)
  struct IMAGEHLP_MODULE64_V2
  {
    DWORD    SizeOfStruct;         // set to sizeof(IMAGEHLP_MODULE64)
    DWORD64  BaseOfImage;          // base load address of module
    DWORD    ImageSize;            // virtual size of the loaded module
    DWORD    TimeDateStamp;        // date/time stamp from pe header
    DWORD    CheckSum;             // checksum from the pe header
    DWORD    NumSyms;              // number of symbols in the symbol table
    SYM_TYPE SymType;              // type of symbols loaded
    TCHAR    ModuleName[32];       // module name
    TCHAR    ImageName[256];       // image name
    TCHAR    LoadedImageName[256]; // symbol file name
  };
  typedef IMAGEHLP_MODULE64_V2  *PIMAGEHLP_MODULE64_V2;

  // since 07-Jun-2002
  struct IMAGEHLP_MODULE64_V3 : IMAGEHLP_MODULE64_V2 
  {
    TCHAR    LoadedPdbName[256];   // pdb file name
    DWORD    CVSig;                // Signature of the CV record in the debug directories
    TCHAR    CVData[MAX_PATH * 3]; // Contents of the CV record
    DWORD    PdbSig;               // Signature of PDB
    GUID     PdbSig70;             // Signature of PDB (VC 7 and up)
    DWORD    PdbAge;               // DBI age of pdb
    BOOL     PdbUnmatched;         // loaded an unmatched pdb
    BOOL     DbgUnmatched;         // loaded an unmatched dbg
    BOOL     LineNumbers;          // we have line number information
    BOOL     GlobalSymbols;        // we have internal symbol information
    BOOL     TypeInfo;             // we have type information
    // new elements: 17-Dec-2003
    BOOL     SourceIndexed;        // pdb supports source server
    BOOL     Publics;              // contains public symbols
  };
  typedef IMAGEHLP_MODULE64_V3  *PIMAGEHLP_MODULE64_V3;

  // reserve enough memory, so the bug in v6.3.5.1 does not lead to memory-overwrites...
  struct T_IMAGEHLP_MODULE64 : IMAGEHLP_MODULE64_V3
  {
    char _padding[4096 - sizeof(IMAGEHLP_MODULE64_V3)];
  };
#pragma pack(pop)

  struct {
    BOOL (WINAPI * Cleanup)(IN HANDLE hProcess);

    PVOID (WINAPI * FunctionTableAccess)(HANDLE hProcess, DWORD64 AddrBase);

    BOOL (WINAPI * GetLineFromAddr)(IN HANDLE hProcess,
                                    IN DWORD64 Address,
                                    OUT PDWORD pdwDisplacement,
                                    OUT T_IMAGEHLP_LINE64 * Line);

    DWORD64 (WINAPI * GetModuleBase)(IN HANDLE hProcess, IN DWORD64 dwAddr);

    BOOL (WINAPI * GetModuleInfo)(IN HANDLE hProcess,
                                  IN DWORD64 Address,
                                  OUT T_IMAGEHLP_MODULE64 * ModuleInfo);

    DWORD (WINAPI * GetOptions)(VOID);

    BOOL (WINAPI * GetSymFromAddr)(IN HANDLE hProcess,
                                   IN DWORD64 Address,
                                   OUT PDWORD64 pdwDisplacement,
                                   OUT PIMAGEHLP_SYMBOL64 Symbol);

    BOOL (WINAPI * FromAddr)(IN HANDLE hProcess,
                             IN DWORD64 Address,
                             OUT PDWORD64 pdwDisplacement,
                             OUT T_SYMBOL_INFO * Symbol);

    BOOL (WINAPI * Initialize)(IN HANDLE hProcess, IN LPCTSTR UserSearchPath, IN BOOL fInvadeProcess);

    DWORD64 (WINAPI * LoadModule)(IN HANDLE hProcess,
                                  IN HANDLE hFile,
                                  IN LPCTSTR ImageName,
                                  IN LPCTSTR ModuleName,
                                  IN DWORD64 BaseOfDll,
                                  IN DWORD SizeOfDll);

    DWORD64 (WINAPI * LoadModuleEx)(IN HANDLE hProcess,
                                    IN HANDLE hFile,
                                    IN LPCTSTR ImageName,
                                    IN LPCTSTR ModuleName,
                                    IN DWORD64 BaseOfDll,
                                    IN DWORD DllSize,
                                    IN PMODLOAD_DATA Data,
                                    IN DWORD Flags);

    DWORD (WINAPI * SetOptions)(IN DWORD SymOptions);

    BOOL (WINAPI * StackWalk)( DWORD                            MachineType,
                               HANDLE                           hProcess,
                               HANDLE                           hThread,
                               LPSTACKFRAME64                   StackFrame,
                               PVOID                            ContextRecord,
                               PREAD_PROCESS_MEMORY_ROUTINE64   ReadMemoryRoutine,
                               PFUNCTION_TABLE_ACCESS_ROUTINE64 FunctionTableAccessRoutine,
                               PGET_MODULE_BASE_ROUTINE64       GetModuleBaseRoutine,
                               PTRANSLATE_ADDRESS_ROUTINE64     TranslateAddress);

    DWORD (WINAPI * UnDecorateName)(LPCTSTR DecoratedName,
                                    LPTSTR  UnDecoratedName,
                                    DWORD   UndecoratedLength,
                                    DWORD   Flags);

    BOOL (WINAPI * GetSearchPath)(HANDLE hProcess, LPTSTR SearchPath, DWORD SearchPathLength);
  } Sym;

  DWORD64 SymLoadModule(HANDLE hProcess, HANDLE hFile, LPCTSTR ImageName, LPCTSTR ModuleName,
                        DWORD64 BaseOfDll, DWORD SizeOfDll) STKWLK_NOEXCEPT
  {
    if (Sym.LoadModuleEx != NULL)
      return Sym.LoadModuleEx(hProcess, hFile, ImageName, ModuleName, BaseOfDll, SizeOfDll, NULL, 0);
    if (Sym.LoadModule != NULL)
      return Sym.LoadModule(hProcess, hFile, ImageName, ModuleName, BaseOfDll, SizeOfDll);
    SetLastError(ERROR_FUNCTION_NOT_CALLED);
    return 0;
  }

  typedef struct _T_SW_SYM_INFO
  {
    union
    {
      IMAGEHLP_SYMBOL64 baseinf;
      T_SYMBOL_INFO     fullinf;
    };
    TCHAR _buffer[StackWalkerBase::STACKWALK_MAX_NAMELEN];
    TCHAR _padding[16];
    enum { Empty, Base, Full } tag;   // header type
  } T_SW_SYM_INFO;

  // return pointer to Symbol.Name
  LPCTSTR SymFromAddr(HANDLE hProcess, DWORD64 Address, PDWORD64 pdwDisplacement, T_SW_SYM_INFO & sym) STKWLK_NOEXCEPT
  {
    if (Sym.FromAddr != NULL)
    {
      memset(&sym.fullinf, 0, sizeof(sym.fullinf));
      sym.fullinf.SizeOfStruct = sizeof(sym.fullinf);
      sym.fullinf.MaxNameLen = _countof(sym._buffer);
      BOOL rc = Sym.FromAddr(hProcess, Address, pdwDisplacement, &sym.fullinf);
      sym.fullinf.Name[(rc == FALSE) ? 0 : sym.fullinf.NameLen] = 0;
      sym.tag = (rc == FALSE) ? T_SW_SYM_INFO::Empty : T_SW_SYM_INFO::Full;
      return (rc == FALSE) ? NULL : sym.fullinf.Name;
    }
#ifndef _UNICODE
    if (Sym.GetSymFromAddr != NULL)
    {
      memset(&sym, 0, sizeof(sym));
      sym.baseinf.SizeOfStruct = sizeof(sym.baseinf);
      sym.baseinf.MaxNameLength = _countof(sym._buffer);
      BOOL rc = Sym.GetSymFromAddr(hProcess, Address, pdwDisplacement, &sym.baseinf);
      sym.tag = (rc == FALSE) ? T_SW_SYM_INFO::Empty : T_SW_SYM_INFO::Base;
      return (rc == FALSE) ? NULL : sym.baseinf.Name;
    }
#endif
    SetLastError(ERROR_FUNCTION_NOT_CALLED);
    return NULL;
  }

private:
// **************************************** ToolHelp32 ************************
#define MAX_MODULE_NAME32 255
#define TH32CS_SNAPMODULE 0x00000008
#pragma pack(push, 8)
  typedef struct tagMODULEENTRY32
  {
    DWORD   dwSize;
    DWORD   th32ModuleID;  // This module
    DWORD   th32ProcessID; // owning process
    DWORD   GlblcntUsage;  // Global usage count on the module
    DWORD   ProccntUsage;  // Module usage count in th32ProcessID's context
    BYTE*   modBaseAddr;   // Base address of module in th32ProcessID's context
    DWORD   modBaseSize;   // Size in bytes of module starting at modBaseAddr
    HMODULE hModule;       // The hModule of this module in th32ProcessID's context
    TCHAR   szModule[MAX_MODULE_NAME32 + 1];
    TCHAR   szExePath[MAX_PATH];
  } MODULEENTRY32;
  typedef MODULEENTRY32* PMODULEENTRY32;
  typedef MODULEENTRY32* LPMODULEENTRY32;
#pragma pack(pop)

  BOOL GetModuleListTH32(HANDLE hProcess, DWORD pid) STKWLK_NOEXCEPT
  {
    // try both dlls...
    LPCWSTR      dllname[] = { L"kernel32.dll", L"tlhelp32.dll" };
    HINSTANCE    hToolhelp = NULL;
    HANDLE (WINAPI * CreateTH32Snapshot)(DWORD dwFlags, DWORD th32ProcessID);
    BOOL (WINAPI * Module32First)(HANDLE hSnapshot, LPMODULEENTRY32 lpme);
    BOOL (WINAPI * Module32Next)(HANDLE hSnapshot, LPMODULEENTRY32 lpme);
    HANDLE        hSnap;
    MODULEENTRY32 me;
    me.dwSize = sizeof(me);

    for (size_t i = 0; i < _countof(dllname); i++)
    {
      hToolhelp = LoadLibraryW(dllname[i]);
      if (hToolhelp == NULL)
        continue;
      int fcnt = 0;
      GetProcAddrEx(fcnt, hToolhelp, "CreateToolhelp32Snapshot", (LPVOID*)&CreateTH32Snapshot);
#ifdef _UNICODE
      GetProcAddrEx(fcnt, hToolhelp, "Module32FirstW", (LPVOID*)&Module32First);
      GetProcAddrEx(fcnt, hToolhelp, "Module32NextW", (LPVOID*)&Module32Next);
#else
      GetProcAddrEx(fcnt, hToolhelp, "Module32First", (LPVOID*)&Module32First);
      GetProcAddrEx(fcnt, hToolhelp, "Module32Next", (LPVOID*)&Module32Next);
#endif
      if (fcnt == 3)
        break; // found the functions!
      FreeLibrary(hToolhelp);
      hToolhelp = NULL;
    }

    if (hToolhelp == NULL)
      return FALSE;

    hSnap = CreateTH32Snapshot(TH32CS_SNAPMODULE, pid);
    if (hSnap == (HANDLE)-1)
    {
      FreeLibrary(hToolhelp);
      return FALSE;
    }

    bool keepGoing = !!Module32First(hSnap, &me);
    int cnt = 0;
    while (keepGoing)
    {
      this->LoadModule(hProcess, me.szExePath, me.szModule, (DWORD64)me.modBaseAddr,
                       me.modBaseSize);
      cnt++;
      keepGoing = !!Module32Next(hSnap, &me);
    }
    CloseHandle(hSnap);
    FreeLibrary(hToolhelp);
    if (cnt <= 0)
      return FALSE;
    return TRUE;
  } // GetModuleListTH32

  // **************************************** PSAPI ************************
  typedef struct _MODULEINFO
  {
    LPVOID lpBaseOfDll;
    DWORD  SizeOfImage;
    LPVOID EntryPoint;
  } MODULEINFO, *LPMODULEINFO;

  typedef struct _SW_MODULE_INFO
  {
    HMODULE   hMods[1024];
    TCHAR     szImgName[2048];
    TCHAR     szModName[2048];
  } SW_MODULE_INFO, *PSW_MODULE_INFO;

  BOOL GetModuleListPSAPI(HANDLE hProcess) STKWLK_NOEXCEPT
  {
    HINSTANCE hPsapi;
    BOOL  (WINAPI * EnumProcessModules)(HANDLE hProcess, HMODULE * lphModule, DWORD cb, LPDWORD lpcbNeeded);
    DWORD (WINAPI * GetModuleFileNameEx)(HANDLE hProcess, HMODULE hModule, LPTSTR lpFilename, DWORD nSize);
    DWORD (WINAPI * GetModuleBaseName)(HANDLE hProcess, HMODULE hModule, LPTSTR lpFilename, DWORD nSize);
    BOOL  (WINAPI * GetModuleInformation)(HANDLE hProcess, HMODULE hModule, LPMODULEINFO pmi, DWORD nSize);

    MODULEINFO       mi;
    SW_MODULE_INFO * mList = NULL;
    DWORD            mListSize = 0;
    size_t           mNumbers;
    size_t           i;
    int              cnt = 0;

    hPsapi = LoadLibraryW(L"psapi.dll");
    if (hPsapi == NULL)
      return FALSE;

    int fcnt = 0;
    GetProcAddrEx(fcnt, hPsapi, "EnumProcessModules", (LPVOID*)&EnumProcessModules);
    GetProcAddrEx(fcnt, hPsapi, "GetModuleInformation", (LPVOID*)&GetModuleInformation);
#ifdef _UNICODE
    GetProcAddrEx(fcnt, hPsapi, "GetModuleFileNameExW", (LPVOID*)&GetModuleFileNameEx);
    GetProcAddrEx(fcnt, hPsapi, "GetModuleBaseNameW", (LPVOID*)&GetModuleBaseName);
#else
    GetProcAddrEx(fcnt, hPsapi, "GetModuleFileNameExA", (LPVOID*)&GetModuleFileNameEx);
    GetProcAddrEx(fcnt, hPsapi, "GetModuleBaseNameA", (LPVOID*)&GetModuleBaseName);
#endif
    if (fcnt < 4)
      goto cleanup;  // we couldn't find all functions

    mList = (SW_MODULE_INFO*) malloc(sizeof(SW_MODULE_INFO));
    if (mList == NULL)
      goto cleanup;

    if (!EnumProcessModules(hProcess, mList->hMods, sizeof(mList->hMods), &mListSize))
      goto cleanup;

    if (mListSize > sizeof(mList->hMods))
      goto cleanup;

    mNumbers = (size_t)mListSize / sizeof(mList->hMods[0]);
    for (i = 0; i < mNumbers; i++)
    {
      HMODULE hMod = mList->hMods[i];
      // base address, size
      GetModuleInformation(hProcess, hMod, &mi, sizeof(mi));
      // image file name
      mList->szImgName[0] = 0;
      GetModuleFileNameEx(hProcess, hMod, mList->szImgName, _countof(mList->szImgName) - 1);
      // module name
      mList->szModName[0] = 0;
      GetModuleBaseName(hProcess, hMod, mList->szModName, _countof(mList->szModName) - 1);

      DWORD dwRes = this->LoadModule(hProcess, mList->szImgName, mList->szModName,
                                    (DWORD64)mi.lpBaseOfDll, mi.SizeOfImage);
      if (dwRes != ERROR_SUCCESS)
        this->m_parent->OnDbgHelpErr(_T("LoadModule"), dwRes, 0);
      cnt++;
    }

  cleanup:
    if (hPsapi != NULL)
      FreeLibrary(hPsapi);
    if (mList != NULL)
      free(mList);

    return cnt != 0;
  } // GetModuleListPSAPI

  DWORD LoadModule(HANDLE hProcess, LPCTSTR img, LPCTSTR mod, DWORD64 baseAddr, DWORD size) STKWLK_NOEXCEPT
  {
    DWORD result = ERROR_SUCCESS;
    ULONGLONG fileVersion = 0;

    if (img == NULL)
      return ERROR_BAD_ARGUMENTS;

    if (this->m_parent == NULL)
      return ERROR_DS_NO_PARENT_OBJECT;

    if (SymLoadModule(hProcess, NULL, img, mod, baseAddr, size) == 0)
    {
      result = GetLastError();
      if (result == ERROR_SUCCESS)
        result = ERROR_DS_SCHEMA_NOT_LOADED;
    }
    {
      // try to retrieve the file-version:
      if ((this->m_parent->m_options & StackWalkerBase::RetrieveFileVersion) != 0)
      {
        GetFileVersion(img, fileVersion);
      }

      // Retrieve some additional-infos about the module
      T_IMAGEHLP_MODULE64 Module;
      LPCTSTR szSymType = _T("-unknown-");
      if (this->GetModuleInfo(hProcess, baseAddr, Module) != FALSE)
      {
        switch (Module.SymType)
        {
          case SymNone:
            szSymType = _T("-nosymbols-");
            break;
          case SymCoff: // 1
            szSymType = _T("COFF");
            break;
          case SymCv: // 2
            szSymType = _T("CV");
            break;
          case SymPdb: // 3
            szSymType = _T("PDB");
            break;
          case SymExport: // 4
            szSymType = _T("-exported-");
            break;
          case SymDeferred: // 5
            szSymType = _T("-deferred-");
            break;
          case SymSym: // 6
            szSymType = _T("SYM");
            break;
          case 7: // SymDia:
            szSymType = _T("DIA");
            break;
          case 8: //SymVirtual:
            szSymType = _T("Virtual");
            break;
        }
      }
      LPCTSTR pdbName = Module.LoadedImageName;
      if (Module.LoadedPdbName[0] != 0)
        pdbName = Module.LoadedPdbName;
      this->m_parent->OnLoadModule(img, mod, baseAddr, size, result, szSymType, pdbName,
                                   fileVersion);
    }
    return result;
  }

public:
  bool GetFileVersion(LPCTSTR filename, ULONGLONG & ver, VS_FIXEDFILEINFO * vinfo = NULL)
  {
    bool result = false;
    BYTE buffer[2048];
    LPVOID vData = (LPVOID)buffer;
    VS_FIXEDFILEINFO * fInfo = NULL;
    ver = 0;
    DWORD dwHandle = 0;
    DWORD dwSize = GetFileVersionInfoSize(filename, &dwHandle);
    if (dwSize == 0 || dwSize > 256 * 1024)
      return false;
    if (dwSize >= sizeof(buffer))
    {
      vData = malloc(dwSize);
      if (vData == NULL)
        return false;
    }
    if (GetFileVersionInfo(filename, dwHandle, dwSize, vData) != 0)
    {
      UINT len = 0;
      BOOL rc = VerQueryValue(vData, _T("\\"), (LPVOID*)&fInfo, &len);
      if (rc != FALSE && fInfo != NULL && len > 0)
      {
        ver = (ULONGLONG)fInfo->dwFileVersionLS + ((ULONGLONG)fInfo->dwFileVersionMS << 32);
        if (vinfo)
          *vinfo = *fInfo;
        result = true;
      }
    }
    if (vData != buffer)
      free(vData);
    return result;
  }

  BOOL LoadModules(HANDLE hProcess, DWORD dwProcessId) STKWLK_NOEXCEPT
  {
    // first try toolhelp32
    if (GetModuleListTH32(hProcess, dwProcessId))
      return true;
    // then try psapi
    return GetModuleListPSAPI(hProcess);
  }

  BOOL GetModuleInfo(HANDLE hProcess, DWORD64 baseAddr, T_IMAGEHLP_MODULE64 & modInfo) STKWLK_NOEXCEPT
  {
    memset(&modInfo, 0, sizeof(modInfo));
    if (Sym.GetModuleInfo == NULL)
    {
      SetLastError(ERROR_DLL_INIT_FAILED);
      return FALSE;
    }
    // First try to use the larger ModuleInfo-Structure
    if (m_IHM64Version == 0 || m_IHM64Version == 3)
    {
      modInfo.SizeOfStruct = sizeof(IMAGEHLP_MODULE64_V3);
      if (Sym.GetModuleInfo(hProcess, baseAddr, &modInfo) != FALSE)
      {
        modInfo.SizeOfStruct = sizeof(IMAGEHLP_MODULE64_V3);
        if (m_IHM64Version == 0)
          m_IHM64Version = 3;
        return TRUE;
      }
      if (m_IHM64Version == 3)
      {
        SetLastError(ERROR_DLL_INIT_FAILED);
        return FALSE;
      }
      if (GetLastError() != ERROR_INVALID_PARAMETER)
        return FALSE;
      // try V2 struct only when SymGetModuleInfo returned error 87
      memset(&modInfo, 0, sizeof(modInfo));
    }

    // could not retrieve the bigger structure, try with the smaller one (as defined in VC7.1)...
    modInfo.SizeOfStruct = sizeof(IMAGEHLP_MODULE64_V2);
    if (Sym.GetModuleInfo(hProcess, baseAddr, &modInfo) != FALSE)
    {
      modInfo.SizeOfStruct = sizeof(IMAGEHLP_MODULE64_V2);
      if (m_IHM64Version == 0)
        m_IHM64Version = 2; // to prevent unnecessary calls with the larger V3 struct...
      return TRUE;
    }
    SetLastError(ERROR_DLL_INIT_FAILED);
    return FALSE;
  }
};

typedef StackWalkerInternal::T_SW_SYM_INFO        T_SW_SYM_INFO;
typedef StackWalkerInternal::T_IMAGEHLP_MODULE64  T_IMAGEHLP_MODULE64;

// #############################################################

#if defined(_MSC_VER) && _MSC_VER >= 1400 && _MSC_VER < 1900
extern "C" void* __cdecl _getptd();
#endif
#if defined(_MSC_VER) && _MSC_VER >= 1900
extern "C" void** __cdecl __current_exception_context();
#endif

static PCONTEXT get_current_exception_context() STKWLK_NOEXCEPT
{
  PCONTEXT * pctx = NULL;
#if defined(_MSC_VER) && _MSC_VER >= 1400 && _MSC_VER < 1900  
  LPSTR ptd = (LPSTR)_getptd();
  if (ptd)
    pctx = (PCONTEXT *)(ptd + (sizeof(void*) == 4 ? 0x8C : 0xF8));
#endif
#if defined(_MSC_VER) && _MSC_VER >= 1900
  pctx = (PCONTEXT *)__current_exception_context();
#endif
  return pctx ? *pctx : NULL;
}

bool StackWalkerBase::Init(ExceptType extype, int options, LPCTSTR szSymPath, DWORD dwProcessId,
                           HANDLE hProcess, PEXCEPTION_POINTERS exp) STKWLK_NOEXCEPT
{
  PCONTEXT ctx = NULL;
  if (extype == AfterCatch)
    ctx = get_current_exception_context();
  if (extype == AfterExcept && exp)
    ctx = exp->ContextRecord;
  this->m_options = options;
  this->m_modulesLoaded = FALSE;
  this->m_szSymPath = NULL;
  this->m_szDbgHelpPath = NULL;
  this->m_MaxRecursionCount = 1000;
  this->m_sw = NULL;
  SetTargetProcess(dwProcessId, hProcess);
  SetSymPath(szSymPath);
  /* MSVC ignore std::nothrow specifier for `new` operator */
  LPVOID buf = malloc(sizeof(StackWalkerInternal));
  if (!buf)
    return false;
  memset(buf, 0, sizeof(StackWalkerInternal));
  this->m_sw = new(buf) StackWalkerInternal(this, this->m_hProcess, ctx);  // placement new
  return true;
}

StackWalkerBase::StackWalkerBase(DWORD dwProcessId, HANDLE hProcess) STKWLK_NOEXCEPT
{
  Init(NonExcept, OptionsAll, NULL, dwProcessId, hProcess);
}

StackWalkerBase::StackWalkerBase(int options, LPCTSTR szSymPath, DWORD dwProcessId, HANDLE hProcess) STKWLK_NOEXCEPT
{
  Init(NonExcept, options, szSymPath, dwProcessId, hProcess);
}

StackWalkerBase::StackWalkerBase(ExceptType extype, int options, PEXCEPTION_POINTERS exp) STKWLK_NOEXCEPT
{
  Init(extype, options, NULL, GetCurrentProcessId(), GetCurrentProcess(), exp);
}

StackWalkerBase::~StackWalkerBase() STKWLK_NOEXCEPT
{
  if (m_sw != NULL) {
    m_sw->~StackWalkerInternal();  // call the object's destructor
    free(m_sw);
  }
  m_sw = NULL;
  SetSymPath(NULL);
  SetDbgHelpPath(NULL);
}

bool StackWalkerBase::SetSymPath(LPCTSTR szSymPath) STKWLK_NOEXCEPT
{
  if (m_szSymPath)
    free(m_szSymPath);
  m_szSymPath = NULL;
  if (szSymPath == NULL)
    return true;
  m_szSymPath = _tcsdup(szSymPath);
  if (m_szSymPath)
    m_options |= SymBuildPath;
  return m_szSymPath ? true : false;
}

bool StackWalkerBase::SetDbgHelpPath(LPCWSTR szDllPath) STKWLK_NOEXCEPT
{
  if (m_szDbgHelpPath)
    free((LPVOID)m_szDbgHelpPath);
  m_szDbgHelpPath = NULL;
  if (szDllPath == NULL)
    return true;
  m_szDbgHelpPath = (LPCWSTR)_wcsdup(szDllPath);
  return m_szDbgHelpPath ? true : false;
}

bool StackWalkerBase::SetTargetProcess(DWORD dwProcessId, HANDLE hProcess) STKWLK_NOEXCEPT
{
  m_dwProcessId = dwProcessId;
  m_hProcess = hProcess;
  if (m_sw)
    m_sw->m_hProcess = hProcess;
  return true;
}

PCONTEXT StackWalkerBase::GetCurrentExceptionContext() STKWLK_NOEXCEPT
{
  return get_current_exception_context();
}

BOOL StackWalkerBase::LoadModules() STKWLK_NOEXCEPT
{
  if (this->m_sw == NULL)
  {
    SetLastError(ERROR_DLL_INIT_FAILED);
    return FALSE;
  }
  if (m_modulesLoaded != FALSE)
    return TRUE;

  // Build the sym-path:
  LPTSTR szSymPath = NULL;
  if ((this->m_options & SymBuildPath) != 0)
  {
    const size_t nSymPathLen = 4096;
    szSymPath = (LPTSTR) malloc((nSymPathLen + 8) * sizeof(TCHAR));
    if (szSymPath == NULL)
    {
      SetLastError(ERROR_NOT_ENOUGH_MEMORY);
      return FALSE;
    }
    szSymPath[0] = 0;
    // Now first add the (optional) provided sympath:
    if (this->m_szSymPath != NULL)
    {
      MyStrCat(szSymPath, nSymPathLen, this->m_szSymPath);
      MyStrCat(szSymPath, nSymPathLen, _T(";"));
    }

    MyStrCat(szSymPath, nSymPathLen, _T(".;"));

    size_t len;
    const size_t nTempLen = 1024;
    TCHAR        szTemp[nTempLen];
    // Now add the current directory:
    len = GetCurrentDirectory(nTempLen, szTemp);
    if (len > 0 && len < nTempLen-1)
    {
      MyStrCat(szSymPath, nSymPathLen, szTemp);
      MyStrCat(szSymPath, nSymPathLen, _T(";"));
    }

    // Now add the path for the main-module:
    len = GetModuleFileName(NULL, szTemp, nTempLen);
    if (len > 0 && len < nTempLen-1)
    {
      for (LPTSTR p = (szTemp + _tcslen(szTemp) - 1); p >= szTemp; --p)
      {
        // locate the rightmost path separator
        if ((*p == '\\') || (*p == '/') || (*p == ':'))
        {
          *p = 0;
          break;
        }
      } // for (search for path separator...)
      if (_tcslen(szTemp) > 0)
      {
        MyStrCat(szSymPath, nSymPathLen, szTemp);
        MyStrCat(szSymPath, nSymPathLen, _T(";"));
      }
    }
    len = GetEnvironmentVariable(_T("_NT_SYMBOL_PATH"), szTemp, nTempLen);
    if (len > 0 && len < nTempLen-1)
    {
      MyStrCat(szSymPath, nSymPathLen, szTemp);
      MyStrCat(szSymPath, nSymPathLen, _T(";"));
    }
    len = GetEnvironmentVariable(_T("_NT_ALTERNATE_SYMBOL_PATH"), szTemp, nTempLen);
    if (len > 0 && len < nTempLen-1)
    {
      MyStrCat(szSymPath, nSymPathLen, szTemp);
      MyStrCat(szSymPath, nSymPathLen, _T(";"));
    }
    len = GetEnvironmentVariable(_T("SYSTEMROOT"), szTemp, nTempLen);
    if (len > 0 && len < nTempLen-1)
    {
      MyStrCat(szSymPath, nSymPathLen, szTemp);
      MyStrCat(szSymPath, nSymPathLen, _T(";"));
      // also add the "system32"-directory:
      MyStrCat(szTemp, nTempLen, _T("\\system32"));
      MyStrCat(szSymPath, nSymPathLen, szTemp);
      MyStrCat(szSymPath, nSymPathLen, _T(";"));
    }

    if ((this->m_options & SymUseSymSrv) != 0)
    {
      LPCTSTR drive = _T("c:\\");
      len = GetEnvironmentVariable(_T("SYSTEMDRIVE"), szTemp, nTempLen);
      if (len > 0 && len < nTempLen-1)
      {
        drive = szTemp;
      }
      MyStrCat(szSymPath, nSymPathLen, _T("SRV*"));
      MyStrCat(szSymPath, nSymPathLen, szTemp);
      MyStrCat(szSymPath, nSymPathLen, _T("\\websymbols*"));
      MyStrCat(szSymPath, nSymPathLen, _T("https://msdl.microsoft.com/download/symbols;"));
    }
  } // if SymBuildPath

  // First Init the whole stuff...
  BOOL bRet = this->m_sw->Init(szSymPath);
  if (szSymPath != NULL)
    free(szSymPath);
  szSymPath = NULL;
  if (bRet == FALSE)
  {
    this->OnDbgHelpErr(_T("Error while initializing dbghelp.dll"), 0, 0);
    SetLastError(ERROR_DLL_INIT_FAILED);
    return FALSE;
  }

  bRet = this->m_sw->LoadModules(this->m_hProcess, this->m_dwProcessId);
  if (bRet != FALSE)
    m_modulesLoaded = TRUE;
  return bRet;
}

// The following is used to pass the "userData"-Pointer to the user-provided readMemoryFunction
// This has to be done due to a problem with the "hProcess"-parameter in x64...
// Because this class is in no case multi-threading-enabled (because of the limitations
// of dbghelp.dll) it is "safe" to use a static-variable
static StackWalkerBase::PReadProcessMemoryRoutine s_readMemoryFunction = NULL;
static LPVOID                                 s_readMemoryFunction_UserData = NULL;

BOOL StackWalkerBase::ShowCallstack(HANDLE                    hThread,
                                const CONTEXT*            context,
                                PReadProcessMemoryRoutine readMemoryFunction,
                                LPVOID                    pUserData) STKWLK_NOEXCEPT
{
  CONTEXT              c;
  CallstackEntry       csEntry;
  T_SW_SYM_INFO        symInf;
  T_IMAGEHLP_MODULE64  Module;
  T_IMAGEHLP_LINE64    Line;
  int                  frameNum;
  bool                 bLastEntryCalled = true;
  int                  curRecursionCount = 0;

  if (this->m_sw == NULL)
  {
    SetLastError(ERROR_OUTOFMEMORY);
    return FALSE;
  }
  if (m_modulesLoaded == FALSE)
    this->LoadModules(); // ignore the result...

  if (this->m_sw->m_hDbhHelp == NULL)
  {
    SetLastError(ERROR_DLL_INIT_FAILED);
    return FALSE;
  }

  s_readMemoryFunction = readMemoryFunction;
  s_readMemoryFunction_UserData = pUserData;

  if (context == NULL)
  {
    // If no context is provided, capture the context
    // See: https://stackwalker.codeplex.com/discussions/446958
#if _WIN32_WINNT <= 0x0501
    // If we need to support XP, we need to use the "old way", because "GetThreadId" is not available!
    if (hThread == GetCurrentThread())
#else
    if (GetThreadId(hThread) == GetCurrentThreadId())
#endif
    {
      memset(&c, 0, sizeof(c));
      if (m_sw->m_ctx.ContextFlags != 0)
        c = m_sw->m_ctx;   // context taken at Init
      else
      {
        c.ContextFlags = STKWLK_CONTEXT_FLAGS;
#if defined(_M_IX86)
        // The following should be enough for walking the callstack...
        __asm    call x
        __asm x: pop eax
        __asm    mov c.Eip, eax
        __asm    mov c.Ebp, ebp
        __asm    mov c.Esp, esp
#else
        // The following is defined for x86 (XP and higher), x64 and IA64
        RtlCaptureContext(&c);
#endif
      }
    }
    else
    {
      SuspendThread(hThread);
      memset(&c, 0, sizeof(CONTEXT));
      c.ContextFlags = STKWLK_CONTEXT_FLAGS;

      // TODO: Detect if you want to get a thread context of a different process, which is running a different processor architecture...
      // This does only work if we are x64 and the target process is x64 or x86;
      // It cannot work, if this process is x64 and the target process is x64... this is not supported...
      // See also: http://www.howzatt.demon.co.uk/articles/DebuggingInWin64.html
      if (GetThreadContext(hThread, &c) == FALSE)
      {
        ResumeThread(hThread);
        return FALSE;
      }
    }
  }
  else
    c = *context;

  // init STACKFRAME for first call
  STACKFRAME64 s; // in/out stackframe
  memset(&s, 0, sizeof(s));
  DWORD imageType;
#ifdef _M_IX86
  // normally, call ImageNtHeader() and use machine info from PE header
  imageType = IMAGE_FILE_MACHINE_I386;
  s.AddrPC.Offset = c.Eip;
  s.AddrPC.Mode = AddrModeFlat;
  s.AddrFrame.Offset = c.Ebp;
  s.AddrFrame.Mode = AddrModeFlat;
  s.AddrStack.Offset = c.Esp;
  s.AddrStack.Mode = AddrModeFlat;
#elif _M_X64
  imageType = IMAGE_FILE_MACHINE_AMD64;
  s.AddrPC.Offset = c.Rip;
  s.AddrPC.Mode = AddrModeFlat;
  s.AddrFrame.Offset = c.Rsp;
  s.AddrFrame.Mode = AddrModeFlat;
  s.AddrStack.Offset = c.Rsp;
  s.AddrStack.Mode = AddrModeFlat;
#elif _M_IA64
  imageType = IMAGE_FILE_MACHINE_IA64;
  s.AddrPC.Offset = c.StIIP;
  s.AddrPC.Mode = AddrModeFlat;
  s.AddrFrame.Offset = c.IntSp;
  s.AddrFrame.Mode = AddrModeFlat;
  s.AddrBStore.Offset = c.RsBSP;
  s.AddrBStore.Mode = AddrModeFlat;
  s.AddrStack.Offset = c.IntSp;
  s.AddrStack.Mode = AddrModeFlat;
#else
#error "Platform not supported!"
#endif

  memset(&Line, 0, sizeof(Line));
  Line.SizeOfStruct = sizeof(Line);

  for (frameNum = 0;; ++frameNum)
  {
    // get next stack frame (StackWalk64(), SymFunctionTableAccess64(), SymGetModuleBase64())
    // if this returns ERROR_INVALID_ADDRESS (487) or ERROR_NOACCESS (998), you can
    // assume that either you are done, or that the stack is so hosed that the next
    // deeper frame could not be found.
    // CONTEXT need not to be supplied if imageTyp is IMAGE_FILE_MACHINE_I386!
    BOOL rc = m_sw->Sym.StackWalk(imageType, m_hProcess, hThread, &s, &c, myReadProcMem, 
                                  m_sw->Sym.FunctionTableAccess, m_sw->Sym.GetModuleBase, NULL);
    if (rc == FALSE)
    {
      // INFO: "StackWalk64" does not set "GetLastError"...
      this->OnDbgHelpErr(_T("StackWalk64"), 0, s.AddrPC.Offset);
      break;
    }

    csEntry.offset = s.AddrPC.Offset;
    csEntry.name[0] = 0;
    csEntry.undName[0] = 0;
    csEntry.undFullName[0] = 0;
    csEntry.offsetFromSmybol = 0;
    csEntry.offsetFromLine = 0;
    csEntry.lineFileName[0] = 0;
    csEntry.lineNumber = 0;
    csEntry.loadedImageName[0] = 0;
    csEntry.moduleName[0] = 0;
    if (s.AddrPC.Offset == s.AddrReturn.Offset)
    {
      if ((this->m_MaxRecursionCount > 0) && (curRecursionCount > m_MaxRecursionCount))
      {
        this->OnDbgHelpErr(_T("StackWalk64-Endless-Callstack!"), 0, s.AddrPC.Offset);
        break;
      }
      curRecursionCount++;
    }
    else
      curRecursionCount = 0;
    if (s.AddrPC.Offset != 0)
    {
      // we seem to have a valid PC
      // show procedure info (SymGetSymFromAddr64())
      LPCTSTR sname = m_sw->SymFromAddr(m_hProcess, s.AddrPC.Offset, &csEntry.offsetFromSmybol, symInf);
      if (sname != NULL)
      {
        MyStrCpy(csEntry.name, STACKWALK_MAX_NAMELEN, sname);
        m_sw->Sym.UnDecorateName(sname, csEntry.undName, STACKWALK_MAX_NAMELEN, UNDNAME_NAME_ONLY);
        m_sw->Sym.UnDecorateName(sname, csEntry.undFullName, STACKWALK_MAX_NAMELEN, UNDNAME_COMPLETE);
      }
      else
      {
        this->OnDbgHelpErr(_T("SymGetSymFromAddr"), GetLastError(), s.AddrPC.Offset);
      }

      // show line number info, NT5.0-method (SymGetLineFromAddr64())
      if (m_sw->Sym.GetLineFromAddr != NULL)
      { // yes, we have SymGetLineFromAddr64()
        BOOL rc = m_sw->Sym.GetLineFromAddr(m_hProcess, s.AddrPC.Offset, &(csEntry.offsetFromLine), &Line);
        if (rc != FALSE)
        {
          csEntry.lineNumber = Line.LineNumber;
          MyStrCpy(csEntry.lineFileName, STACKWALK_MAX_NAMELEN, Line.FileName);
        }
        else
        {
          this->OnDbgHelpErr(_T("SymGetLineFromAddr64"), GetLastError(), s.AddrPC.Offset);
        }
      } // yes, we have SymGetLineFromAddr64()

      // show module info (SymGetModuleInfo64())
      if (this->m_sw->GetModuleInfo(this->m_hProcess, s.AddrPC.Offset, Module) != FALSE)
      { // got module info OK
        switch (Module.SymType)
        {
          case SymNone:
            csEntry.symTypeString = _T("-nosymbols-");
            break;
          case SymCoff:
            csEntry.symTypeString = _T("COFF");
            break;
          case SymCv:
            csEntry.symTypeString = _T("CV");
            break;
          case SymPdb:
            csEntry.symTypeString = _T("PDB");
            break;
          case SymExport:
            csEntry.symTypeString = _T("-exported-");
            break;
          case SymDeferred:
            csEntry.symTypeString = _T("-deferred-");
            break;
          case SymSym:
            csEntry.symTypeString = _T("SYM");
            break;
#if API_VERSION_NUMBER >= 9
          case SymDia:
            csEntry.symTypeString = _T("DIA");
            break;
#endif
          case 8: //SymVirtual:
            csEntry.symTypeString = _T("Virtual");
            break;
          default:
            //_snprintf( ty, sizeof(ty), "symtype=%ld", (long) Module.SymType );
            csEntry.symTypeString = NULL;
            break;
        }

        MyStrCpy(csEntry.moduleName, STACKWALK_MAX_NAMELEN, Module.ModuleName);
        csEntry.baseOfImage = Module.BaseOfImage;
        MyStrCpy(csEntry.loadedImageName, STACKWALK_MAX_NAMELEN, Module.LoadedImageName);
      } // got module info OK
      else
      {
        this->OnDbgHelpErr(_T("SymGetModuleInfo64"), GetLastError(), s.AddrPC.Offset);
      }
    } // we seem to have a valid PC

    CallstackEntryType et = nextEntry;
    if (frameNum == 0)
      et = firstEntry;
    bLastEntryCalled = false;
    this->OnCallstackEntry(et, csEntry);

    if (s.AddrReturn.Offset == 0)
    {
      bLastEntryCalled = true;
      this->OnCallstackEntry(lastEntry, csEntry);
      SetLastError(ERROR_SUCCESS);
      break;
    }
  } // for ( frameNum )

  if (bLastEntryCalled == false)
    this->OnCallstackEntry(lastEntry, csEntry);

  if (context == NULL)
    ResumeThread(hThread);

  return TRUE;
}

BOOL StackWalkerBase::ShowCallstack(const CONTEXT * context) STKWLK_NOEXCEPT
{
  return ShowCallstack(GetCurrentThread(), context, NULL, NULL);
}

BOOL StackWalkerBase::ShowObject(LPVOID pObject) STKWLK_NOEXCEPT
{
  BOOL result = FALSE;
  LPCTSTR sname = NULL;
  if (this->m_sw == NULL)
  {
    SetLastError(ERROR_OUTOFMEMORY);
    goto fin;
  }
  // Load modules if not done yet
  if (m_modulesLoaded == FALSE)
    this->LoadModules(); // ignore the result...

  // Verify that the DebugHelp.dll was actually found
  if (this->m_sw->m_hDbhHelp == NULL)
  {
    SetLastError(ERROR_DLL_INIT_FAILED);
    goto fin;
  }

  // SymGetSymFromAddr64 or SymFromAddr is required
  if (m_sw->Sym.GetSymFromAddr == NULL && m_sw->Sym.FromAddr == NULL)
    goto fin;

  {
    // Show object info (SymGetSymFromAddr64())
    T_SW_SYM_INFO symInf;
    DWORD64 dwAddress = (DWORD64)pObject;
    DWORD64 dwDisplacement = 0;
    sname = m_sw->SymFromAddr(m_hProcess, dwAddress, &dwDisplacement, symInf);
    if (sname == NULL)
      this->OnDbgHelpErr(_T("SymGetSymFromAddr"), GetLastError(), dwAddress);
    else
      result = TRUE;
  }
fin:  
  // Object name output
  this->OnShowObject(pObject, sname);
  return result;
};

BOOL WINAPI StackWalkerBase::myReadProcMem(HANDLE  hProcess,
                                           DWORD64 qwBaseAddress,
                                           PVOID   lpBuffer,
                                           DWORD   nSize,
                                           LPDWORD lpNumberOfBytesRead) STKWLK_NOEXCEPT
{
  if (s_readMemoryFunction == NULL)
  {
    SIZE_T st;
    BOOL   bRet = ReadProcessMemory(hProcess, (LPVOID)qwBaseAddress, lpBuffer, nSize, &st);
    *lpNumberOfBytesRead = (DWORD)st;
    //printf("ReadMemory: hProcess: %p, baseAddr: %p, buffer: %p, size: %d, read: %d, result: %d\n", hProcess, (LPVOID) qwBaseAddress, lpBuffer, nSize, (DWORD) st, (DWORD) bRet);
    return bRet;
  }
  else
  {
    return s_readMemoryFunction(hProcess, qwBaseAddress, lpBuffer, nSize, lpNumberOfBytesRead,
                                s_readMemoryFunction_UserData);
  }
}

// =====================================================================================

void StackWalkerDemo::OnLoadModule(LPCTSTR   img,
                               LPCTSTR   mod,
                               DWORD64   baseAddr,
                               DWORD     size,
                               DWORD     result,
                               LPCTSTR   symType,
                               LPCTSTR   pdbName,
                               ULONGLONG fileVersion) STKWLK_NOEXCEPT
{
  TCHAR buf[STACKWALK_MAX_NAMELEN];
  if (fileVersion == 0)
    MyTStrFmt(buf, _countof(buf), _T("%s:%s (%p), size: %d (result: %d), SymType: '%s', PDB: '%s'\n"),
              img, mod, (LPVOID)baseAddr, size, result, symType, pdbName);
  else
  {
    DWORD v4 = (DWORD)(fileVersion & 0xFFFF);
    DWORD v3 = (DWORD)((fileVersion >> 16) & 0xFFFF);
    DWORD v2 = (DWORD)((fileVersion >> 32) & 0xFFFF);
    DWORD v1 = (DWORD)((fileVersion >> 48) & 0xFFFF);
    MyTStrFmt(
        buf, _countof(buf),
        _T("%s:%s (%p), size: %d (result: %d), SymType: '%s', PDB: '%s', fileVersion: %d.%d.%d.%d\n"),
        img, mod, (LPVOID)baseAddr, size, result, symType, pdbName, v1, v2, v3, v4);
  }
  OnOutput(buf);
}

void StackWalkerDemo::OnCallstackEntry(CallstackEntryType eType, CallstackEntry& entry) STKWLK_NOEXCEPT
{
  TCHAR buf[STACKWALK_MAX_NAMELEN];
  if ((eType != lastEntry) && (entry.offset != 0))
  {
    if (entry.name[0] == 0)
      MyStrCpy(entry.name, _countof(entry.name), _T("(function-name not available)"));
    if (entry.undName[0] != 0)
      MyStrCpy(entry.name, _countof(entry.name), entry.undName);
    if (entry.undFullName[0] != 0)
      MyStrCpy(entry.name, _countof(entry.name), entry.undFullName);
    if (entry.lineFileName[0] == 0)
    {
      MyStrCpy(entry.lineFileName, _countof(entry.lineFileName), _T("(filename not available)"));
      if (entry.moduleName[0] == 0)
        MyStrCpy(entry.moduleName, _countof(entry.moduleName), _T("(module-name not available)"));
      MyTStrFmt(buf, _countof(buf), _T("%p (%s): %s: %s\n"),
                (LPVOID)entry.offset, entry.moduleName, entry.lineFileName, entry.name);
    }
    else
      MyTStrFmt(buf, _countof(buf), _T("%s (%d): %s\n"),
                entry.lineFileName, entry.lineNumber, entry.name);
    OnOutput(buf);
  }
}

void StackWalkerDemo::OnShowObject(LPVOID pObject, LPCTSTR szName) STKWLK_NOEXCEPT
{
  TCHAR buf[STACKWALK_MAX_NAMELEN];
  MyTStrFmt(buf, _countof(buf), _T("Object: Addr: %p, Name: \"%s\"\n"), pObject, szName);
  OnOutput(buf);
}

void StackWalkerDemo::OnDbgHelpErr(LPCTSTR szFuncName, DWORD gle, DWORD64 addr) STKWLK_NOEXCEPT
{
  TCHAR buf[STACKWALK_MAX_NAMELEN];
  MyTStrFmt(buf, _countof(buf), _T("ERROR: %s, GetLastError: %d (Address: %p)\n"),
            szFuncName, gle, (LPVOID)addr);
  OnOutput(buf);
}

void StackWalkerDemo::OnLoadDbgHelp(ULONGLONG verFile, LPCTSTR szDllPath) STKWLK_NOEXCEPT
{
  TCHAR buf[STACKWALK_MAX_NAMELEN];
  DWORD v4 = (DWORD)(verFile & 0xFFFF);
  DWORD v3 = (DWORD)((verFile >> 16) & 0xFFFF);
  DWORD v2 = (DWORD)((verFile >> 32) & 0xFFFF);
  DWORD v1 = (DWORD)((verFile >> 48) & 0xFFFF);
  MyTStrFmt(buf, _countof(buf), _T("LoadDbgHelp: FileVer: %d.%d.%d.%d, Path: \"%s\"\n"),
            v1, v2, v3, v4, szDllPath);
  OnOutput(buf);
}

void StackWalkerDemo::OnSymInit(LPCTSTR szSearchPath, DWORD symOptions, LPCTSTR szUserName) STKWLK_NOEXCEPT
{
  TCHAR buf[STACKWALK_MAX_NAMELEN];
  MyTStrFmt(buf, _countof(buf), _T("SymInit: symOptions: 0x%08X, UserName: \"%s\"\n"),
            symOptions, szUserName);
  OnOutput(buf);

  MyTStrFmt(buf, _countof(buf), _T("Symbol-SearchPath: \"%s\"\n"), szSearchPath);
  OnOutput(buf);

  // Also display the OS-version
  OSVERSIONINFOEX ver = { 0 };
  ver.dwOSVersionInfoSize = sizeof(ver);
#if _MSC_VER >= 1900
#pragma warning(push)
#pragma warning(disable : 4996)  // For fix warning "GetVersionExW was declared deprecated"
#endif
  if (GetVersionEx((OSVERSIONINFO*)&ver) != FALSE)
  {
    MyTStrFmt(buf, _countof(buf), _T("OS-Version: %d.%d.%d (%s) 0x%04x-%d\n"),
              ver.dwMajorVersion, ver.dwMinorVersion, ver.dwBuildNumber, ver.szCSDVersion,
              ver.wSuiteMask, ver.wProductType);
    OnOutput(buf);
  }
#if _MSC_VER >= 1900
#pragma warning(pop)
#endif
}

void StackWalkerDemo::OnOutput(LPCTSTR buffer) STKWLK_NOEXCEPT
{
  OutputDebugString(buffer);
}
