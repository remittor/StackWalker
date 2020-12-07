#ifndef __STACKWALKER_H__
#define __STACKWALKER_H__

#if defined(_MSC_VER)

/**********************************************************************
 *
 * StackWalker.h
 *
 *
 *
 * LICENSE (http://www.opensource.org/licenses/bsd-license.php)
 *
 *   Copyright (c) 2005-2009, Jochen Kalmbach
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
 * **********************************************************************/
// #pragma once is supported starting with _MSC_VER 1000,
// so we need not to check the version (because we only support _MSC_VER >= 1100)!
#pragma once

#include <windows.h>
#include <tchar.h>

// special defines for VC5/6 (if no actual PSDK is installed):
#if _MSC_VER < 1300
typedef unsigned __int64 DWORD64, *PDWORD64;
#if defined(_WIN64)
typedef unsigned __int64 SIZE_T, *PSIZE_T;
#else
typedef unsigned long SIZE_T, *PSIZE_T;
#endif
#endif // _MSC_VER < 1300

#ifndef STKWLK_THROWABLE
#if _MSC_VER < 1900    /* noexcept added since vs2015 */
#define STKWLK_NOEXCEPT throw()
#else
#define STKWLK_NOEXCEPT noexcept
#endif
#endif // STKWLK_THROWABLE

#if _MSC_VER >= 1800
#define STKWLK_DEFAULT  =default
#define STKWLK_DELETED  =delete
#else
#define STKWLK_DEFAULT 
#define STKWLK_DELETED 
#endif

#if _MSC_VER >= 1700
#define STKWLK_FINAL final
#elif _MSC_VER >= 1400
#define STKWLK_FINAL sealed
#else
#define STKWLK_FINAL
#endif


class StackWalkerInternal; // forward

class StackWalkerBase
{
public:
  typedef enum ExceptType
  {
    NonExcept   = 0,     // RtlCaptureContext
    AfterExcept = 1,
    AfterCatch  = 2,     // get_current_exception_context
  } ExceptType;

  typedef enum StackWalkOptions
  {
    // No addition info will be retrieved
    // (only the address is available)
    RetrieveNone = 0,

    // Try to get the symbol-name
    RetrieveSymbol = 1,

    // Try to get the line for this symbol
    RetrieveLine = 2,

    // Try to retrieve the module-infos
    RetrieveModuleInfo = 4,

    // Also retrieve the version for the DLL/EXE
    RetrieveFileVersion = 8,

    // Contains all the above
    RetrieveVerbose = 0xF,

    // Generate a "good" symbol-search-path
    SymBuildPath = 0x10,

    // Also use the public Microsoft-Symbol-Server
    SymUseSymSrv = 0x20,

    // Contains all the above "Sym"-options
    SymAll = 0x30,

    // Contains all options (default)
    OptionsAll = 0x3F
  } StackWalkOptions;

  StackWalkerBase(ExceptType extype,
                  int options = OptionsAll,
                  PEXCEPTION_POINTERS exp = NULL) STKWLK_NOEXCEPT;

  StackWalkerBase(int     options = OptionsAll, // 'int' is by design, to combine the enum-flags
                  LPCTSTR szSymPath = NULL,
                  DWORD   dwProcessId = GetCurrentProcessId(),
                  HANDLE  hProcess = GetCurrentProcess()) STKWLK_NOEXCEPT;

  StackWalkerBase(DWORD dwProcessId, HANDLE hProcess) STKWLK_NOEXCEPT;

  // delete copy constructor
  StackWalkerBase(const StackWalkerBase & ) STKWLK_DELETED;
  const StackWalkerBase & operator = ( const StackWalkerBase & ) STKWLK_DELETED;
#if _MSC_VER >= 1800
  // delete move constructor
  StackWalkerBase(StackWalkerBase &&) STKWLK_DELETED;
  StackWalkerBase & operator = (StackWalkerBase && ) STKWLK_DELETED;
#endif

  virtual ~StackWalkerBase() STKWLK_NOEXCEPT;

  bool SetSymPath(LPCTSTR szSymPath) STKWLK_NOEXCEPT;
  
  bool SetDbgHelpPath(LPCWSTR szDllPath) STKWLK_NOEXCEPT;

  bool SetTargetProcess(DWORD dwProcessId, HANDLE hProcess) STKWLK_NOEXCEPT;

  PCONTEXT GetCurrentExceptionContext() STKWLK_NOEXCEPT;

  LPVOID GetUserData() STKWLK_NOEXCEPT;

private:
  bool Init(ExceptType extype, int options, LPCTSTR szSymPath, DWORD dwProcessId,
            HANDLE hProcess, PEXCEPTION_POINTERS exp = NULL) STKWLK_NOEXCEPT;

public:
  typedef BOOL (WINAPI * PReadMemRoutine)(
      HANDLE  hProcess,
      DWORD64 qwBaseAddress,
      PVOID   lpBuffer,
      DWORD   nSize,
      LPDWORD lpNumberOfBytesRead,
      LPVOID  pUserData // optional data, which was passed in "ShowCallstack"
  );

  bool ShowModules(LPVOID pUserData = NULL) STKWLK_NOEXCEPT;

  bool ShowCallstack(const CONTEXT * context, LPVOID pUserData = NULL) STKWLK_NOEXCEPT;

  bool ShowCallstack(HANDLE          hThread = GetCurrentThread(),
                     const CONTEXT * context = NULL,
                     PReadMemRoutine pReadMemFunc = NULL,
                     LPVOID          pUserData = NULL) STKWLK_NOEXCEPT;

  bool ShowObject(LPVOID pObject, LPVOID pUserData = NULL) STKWLK_NOEXCEPT;

protected:
  struct TFileVer
  {
    WORD  wMajor;
    WORD  wMinor;
    WORD  wBuild;
    WORD  wRevis;
    TFileVer() STKWLK_NOEXCEPT { zeroinit(); }
    void zeroinit() STKWLK_NOEXCEPT { wMajor = 0; wMinor = 0; wBuild = 0; wRevis = 0; }
    bool isEmpty() const { return !wMajor && !wMinor && !wBuild && !wRevis; }
  };

  struct TLoadDbgHelp
  {
    TFileVer ver;
    LPCTSTR  szDllPath;
  };
  virtual void OnLoadDbgHelp(const TLoadDbgHelp & data) STKWLK_NOEXCEPT = 0;

  struct TSymInit
  {
    LPCTSTR  szSearchPath;
    DWORD    dwSymOptions;
    LPCTSTR  szUserName;
  };
  virtual void OnSymInit(const TSymInit & data) STKWLK_NOEXCEPT = 0;

  struct TLoadModule
  {
    LPCTSTR  imgName;
    LPCTSTR  modName;
    DWORD64  baseAddr;
    DWORD    size;
    DWORD    result;
    LPCTSTR  symType;
    LPCTSTR  pdbName;
    TFileVer ver;
  };
  virtual void OnLoadModule(const TLoadModule & data) STKWLK_NOEXCEPT = 0;

  enum CallstackEntryType
  {
    firstEntry,
    nextEntry,
    lastEntry
  };
  struct TCallstackEntry      // Entry for each Callstack-Entry
  {
    CallstackEntryType type;
    DWORD64  offset;           // if 0, we have no valid entry
    LPCTSTR  name;
    LPCTSTR  undName;
    LPCTSTR  undFullName;
    DWORD64  offsetFromSymbol;
    DWORD    offsetFromLine;
    DWORD    lineNumber;
    LPCTSTR  lineFileName;
    DWORD    symType;
    LPCTSTR  symTypeString;
    LPCTSTR  moduleName;
    DWORD64  baseOfImage;
    LPCTSTR  loadedImageName;
  };
  virtual void OnCallstackEntry(const TCallstackEntry & entry) STKWLK_NOEXCEPT = 0;

  struct TShowObject
  {
    LPVOID   pObject;
    LPCTSTR  szName;
  };
  virtual void OnShowObject(const TShowObject & data) STKWLK_NOEXCEPT = 0;

  struct TDbgHelpErr
  {
    LPCTSTR  szFuncName;
    DWORD    gle;
    DWORD64  addr;
    TDbgHelpErr(LPCTSTR szFuncName, DWORD gle = 0, DWORD64 addr = 0) STKWLK_NOEXCEPT;
  };
  virtual void OnDbgHelpErr(const TDbgHelpErr & data) STKWLK_NOEXCEPT = 0;

  StackWalkerInternal* m_sw;
  HANDLE               m_hProcess;
  DWORD                m_dwProcessId;
  LPTSTR               m_szSymPath;
  LPCWSTR              m_szDbgHelpPath;

  int m_options;
  int m_MaxRecursionCount;

  friend StackWalkerInternal;
}; // class StackWalkerBase


class StackWalkerDemo : public StackWalkerBase
{
public:  
  StackWalkerDemo(ExceptType extype,
                  int options = OptionsAll,
                  PEXCEPTION_POINTERS exp = NULL) STKWLK_NOEXCEPT
    : StackWalkerBase(extype, options, exp)
  { }

  StackWalkerDemo(int     options = OptionsAll, // 'int' is by design, to combine the enum-flags
                  LPCTSTR szSymPath = NULL,
                  DWORD   dwProcessId = GetCurrentProcessId(),
                  HANDLE  hProcess = GetCurrentProcess()) STKWLK_NOEXCEPT
    : StackWalkerBase(options, szSymPath, dwProcessId, hProcess)
  { }

  StackWalkerDemo(DWORD dwProcessId, HANDLE hProcess) STKWLK_NOEXCEPT
    : StackWalkerBase(dwProcessId, hProcess)
  { }

  virtual void OnLoadDbgHelp(const TLoadDbgHelp & data) STKWLK_NOEXCEPT;

  virtual void OnSymInit(const TSymInit & data) STKWLK_NOEXCEPT;

  virtual void OnLoadModule(const TLoadModule & data) STKWLK_NOEXCEPT;

  virtual void OnCallstackEntry(const TCallstackEntry & entry) STKWLK_NOEXCEPT;

  virtual void OnShowObject(const TShowObject & data) STKWLK_NOEXCEPT;

  virtual void OnDbgHelpErr(const TDbgHelpErr & data) STKWLK_NOEXCEPT;

  virtual void OnOutput(LPCTSTR szText) STKWLK_NOEXCEPT;

}; // class StackWalkerDemo


#endif //defined(_MSC_VER)

#endif // __STACKWALKER_H__
