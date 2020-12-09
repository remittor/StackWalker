# StackWalker - Walking the callstack

The original project is located here: https://github.com/JochenKalmbach/StackWalker

This article describes the (documented) way to walk a callstack for any thread (own, other and remote). It has an abstraction layer, so the calling app does not need to know the internals.

# Latest Build status

[![build result](https://ci.appveyor.com/api/projects/status/github/remittor/StackWalker?branch=master&svg=true)](https://ci.appveyor.com/project/remittor/StackWalker)

# Documentation

## Introduction

In some cases you need to display the callstack of the current thread or you are just interested in the callstack of other threads. This project was created for this very reason.

The goal for this project was the following:

* Simple interface to generate a callstack
* C++ based to allow overwrites of several methods
* Hiding the implementation details (API) from the class interface
* Support of x86 and x86-64 architectures
* Default output to debugger-output window (but can be customized)
* Support of user-provided read-memory-function
* Most portable solution to walk the callstack for MSVC
* Support multi-threading usage
* Support unicode file names
* Support C++ exception

Minimum requirements: WinXP

## Background

To walk the callstack there is a documented interface: [StackWalk64](http://msdn.microsoft.com/library/en-us/debug/base/stackwalk64.asp).

The latest *dbghelp.dll* can always be downloaded with the [Debugging Tools for Windows](http://www.microsoft.com/whdc/devtools/debugging/).

Alternative download page: http://eretik.omegahg.com/kd/WinDBG_Download.html

## Build

```
mkdir build-dir
cd build-dir

# batch
cmake -G "Visual Studio 15 2017 Win64" --config RelWithDebInfo -DCMAKE_INSTALL_PREFIX=%cd%/root ..
# powershell
cmake -G "Visual Studio 15 2017 Win64" --config RelWithDebInfo -DCMAKE_INSTALL_PREFIX="$($(get-location).Path)/root" ..

cmake --build . --config RelWithDebInfo
ctest.exe -V -C RelWithDebInfo
cmake --build . --target install --config RelWithDebInfo
```

## Using the code

The usage of the class is very simple. For example if you want to display the callstack of the current thread, just instantiate a `StackWalkDemo` object and call the `ShowCallstack` member:

```c++
#include "StackWalker.h"
#pragma optimize( "", off )

void Func4()
{
    StackWalkerDemo sw;
    sw.ShowCallstack();
}
void Func3()
{
    Func4();
}
void Func2()
{
    Func3();
}
void Func1()
{
    Func2();
}

int main()
{
    Func1();
    return 0;
}
```

This produces the following output in the debugger-output window:
```
d:\dev\stackwalker\src\stackwalker.cpp (1760): StackWalkerBase::ShowCallstack
d:\dev\stackwalker\test\main.cpp (7): Func4
d:\dev\stackwalker\test\main.cpp (12): Func3
d:\dev\stackwalker\test\main.cpp (16): Func2
d:\dev\stackwalker\test\main.cpp (20): Func1
d:\dev\stackwalker\test\main.cpp (25): main
f:\dd\vctools\crt\vcstartup\src\startup\exe_common.inl (253): __scrt_common_main_seh
00000000776E556D (kernel32): (filename not available): BaseThreadInitThunk
000000007794372D (ntdll): (filename not available): RtlUserThreadStart
```

### Providing own output-mechanism

If you want to direct the output to a file or want to use some other output-mechanism, you simply need to derive from the `StackWalkerBase` or `StackWalkerDemo` class.
If you are satisfied with the output format used in class `StackWalkerDemo`, then you should derive from this class and overwrite method `OnOutput`.
To output also to the console, you need to do the following:
```c++
class StackWalker : public StackWalkerDemo
{
public:
    StackWalker() : StackWalkerDemo() { }

protected:
    virtual void OnOutput(LPCWSTR szText) STKWLK_NOEXCEPT
    {
        wprintf(szText);                    // out text to console
        StackWalkerDemo::OnOutput(szText);  // out text to DebugLog
    }
};
```

### Retrieving detailed callstack info

If you want detailed info about the callstack (like loaded-modules, addresses, errors, ...) you can overwrite the corresponding methods. The following methods are provided:
```c++
class StackWalkerBase
{
protected:
    virtual void OnLoadDbgHelp   (const TLoadDbgHelp & data)     noexcept = 0;
    virtual void OnSymInit       (const TSymInit & data)         noexcept = 0;
    virtual void OnLoadModule    (const TLoadModule & data)      noexcept = 0;
    virtual void OnCallstackEntry(const TCallstackEntry & entry) noexcept = 0;
    virtual void OnShowObject    (const TShowObject & data)      noexcept = 0;
    virtual void OnDbgHelpErr    (const TDbgHelpErr & data)      noexcept = 0;    
};
```

These methods are called during the generation of the callstack.

### Various kinds of callstacks

In the constructor of the class, you need to specify if you want to generate callstacks for the current process or for another process. The following constructors are available:
```c++
class StackWalkerBase
{
public:
    // For showing stack trace after __except or catch
    StackWalkerBase(ExceptType extype,     // possible values: AfterExcept or AfterCatch
                    int options = OptionsAll,
                    PEXCEPTION_POINTERS exp = NULL) noexcept;

    StackWalkerBase(int     options = OptionsAll,
                    SW_CSTR szSymPath = NULL,
                    DWORD   dwProcessId = GetCurrentProcessId(),
                    HANDLE  hProcess = GetCurrentProcess()) noexcept;
    
    // Just for other processes with default-values for options and symPath
    StackWalkerBase(DWORD dwProcessId, HANDLE hProcess) noexcept;
};
```

To do the actual stack-walking you need to call any of the following functions:
```c++
class StackWalkerBase
{
public:
    bool ShowCallstack(const CONTEXT * context,
                       LPVOID          pUserData = NULL) noexcept;

    bool ShowCallstack(HANDLE          hThread = GetCurrentThread(),
                       const CONTEXT * context = NULL,
                       PReadMemRoutine pReadMemFunc = NULL,
                       LPVOID          pUserData = NULL) noexcept;
};
```

### Displaying the callstack of an exception

With this `StackWalker` you can also display the callstack inside an structured exception handler. You only need to write a filter-function which does the stack-walking:
```c++
// The exception filter function:
LONG WINAPI ExpFilter(EXCEPTION_POINTERS * pExp, DWORD dwExpCode)
{
    StackWalkerDemo sw;
    sw.ShowCallstack(GetCurrentThread(), pExp->ContextRecord);
    return EXCEPTION_EXECUTE_HANDLER;
}

// This is how to catch an exception:
__try
{
    // do some ugly stuff...
}
__except (ExpFilter(GetExceptionInformation(), GetExceptionCode()))
{
}
```

Display the callstack inside an C++ exception handler (two ways):
```c++
// This is how to catch an exception:
try
{
    // do some ugly stuff...
}
catch (std::exception & ex)
{
    StackWalkerDemo sw;
    sw.ShowCallstack(GetCurrentThread(), sw.GetCurrentExceptionContext());
}
catch (...)
{
    StackWalkerDemo sw(StackWalker::AfterCatch);
    sw.ShowCallstack();
}
```

### Options

To do some kind of modification of the behavior, you can optionally specify some options. Here is the list of the available options:

```c++
typedef enum StackWalkOptions
{
    // No addition info will be retrieved (only the address is available)
    RetrieveNone = 0,

    // Try to get the symbol-name
    RetrieveSymbol = 1,

    // Try to get the line for this symbol
    RetrieveLine = 2,

    // Try to retrieve the module-infos
    RetrieveModuleInfo = 4,

    // Also retrieve the version for the DLL/EXE
    RetrieveFileVersion = 8,

    // Generate a "good" symbol-search-path
    SymBuildPath = 0x10,

    // Also use the public Microsoft-Symbol-Server
    SymUseSymSrv = 0x20,

    // DbgHelp.DLL will be loaded once in a separate place in the process memory
    SymIsolated = 0x40,

} StackWalkOptions;

// Contains all the "Retrieve"-options
enum { RetrieveVerbose = RetrieveSymbol | RetrieveLine | RetrieveModuleInfo | RetrieveFileVersion };

// Contains all the "Sym"-options (excluding `SymIsolated`)
enum { SymAll = SymBuildPath | SymUseSymSrv };

// Contains all options (default)
enum { OptionsAll = RetrieveVerbose | SymAll };
```
