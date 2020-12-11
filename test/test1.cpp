#if defined(STKWLK_UNIT_TEST) && STKWLK_UNIT_TEST == 1 

#if defined(STKWLK_ANSI) || defined(_MBCS)
#error "Support only unicode"
#endif

#include "StackWalker.h"
#include <stdio.h>
#include <vector>

#pragma optimize( "", off )

bool g_EHasync = false;   // true if compiled with option /EHa

// =========================================================================================

typedef int (* FnTestCallstackEntry)(const StackWalkerBase::TCallstackEntry & entry);

struct TestContext
{
  struct CallData
  {
    LPVOID   addr;
    LPCWSTR  name;
    int      line;
    CallData(LPVOID addr, LPCWSTR name = NULL, int line = 0);
  };
  int  m_level;
  bool m_callParent;
  FnTestCallstackEntry  m_testCallstackEntry;
  std::vector<CallData> m_callList;

  TestContext();
  void reset(int level = -1, FnTestCallstackEntry cb = NULL);
  void AddCall(LPVOID addr, LPCWSTR name = NULL, int line = 0);
  bool UpdateLevel(LPCWSTR name, LPCWSTR target, bool skip_target = true);
  int  CheckEntry(const StackWalkerBase::TCallstackEntry & entry);
};

class StackWalker : public StackWalkerDemo
{
public:
  enum { OptionsAll = StackWalkerBase::RetrieveVerbose | StackWalkerBase::SymBuildPath };

  StackWalker() STKWLK_NOEXCEPT
    : StackWalkerDemo(OptionsAll)
  { 
    // nothing
  }

  StackWalker(ExceptType extype, PEXCEPTION_POINTERS exp = NULL) STKWLK_NOEXCEPT
    : StackWalkerDemo(extype, OptionsAll, exp)
  {
    // nothing
  }

  bool ShowCallstack(const CONTEXT * context, LPVOID pUserData = NULL) STKWLK_NOEXCEPT
  {
    OnOutput(L"________CALLSTACK________:\n");
    return StackWalkerDemo::ShowCallstack(context, pUserData);
  }

  bool ShowCallstack(HANDLE          hThread = GetCurrentThread(),
                     const CONTEXT * context = NULL,
                     PReadMemRoutine pReadMemFunc = NULL,
                     LPVOID          pUserData = NULL) STKWLK_NOEXCEPT
  {
    OnOutput(L"________callstack________:\n");
    return StackWalkerDemo::ShowCallstack(hThread, context, pReadMemFunc, pUserData);
  }

  virtual void OnSymInit(const TSymInit & data) STKWLK_NOEXCEPT
  {
    // nothing
  }

  virtual void OnLoadModule(const TLoadModule & data) STKWLK_NOEXCEPT
  {
    // nothing
  }

  virtual void OnCallstackEntry(const TCallstackEntry & entry) STKWLK_NOEXCEPT
  {
    TestContext * ctx = (TestContext *)GetUserData();
    if (!ctx || ctx->m_callParent)
      StackWalkerDemo::OnCallstackEntry(entry);
    if (ctx && ctx->m_testCallstackEntry)
      ctx->m_testCallstackEntry(entry);
  }

  virtual void OnOutput(LPCWSTR szText) STKWLK_NOEXCEPT
  {
    wprintf(L"%s", szText);
  }
};

// =========================================================================================

TestContext::CallData::CallData(LPVOID addr, LPCWSTR name, int line)
{ 
  this->addr = addr;
  this->name = name;
  this->line = line;
}

TestContext::TestContext()
{
  reset();
}

void TestContext::reset(int level, FnTestCallstackEntry cb)
{
  m_level = level;
  m_callParent = cb ? true : false;
  m_testCallstackEntry = cb;
  m_callList.reserve(256);
  m_callList.clear();
}

bool TestContext::UpdateLevel(LPCWSTR name, LPCWSTR target, bool skip_target)
{
  if (m_level < 0) {
    LPCWSTR n = wcsstr(name, target);
    if (!n && m_level < -1)
      return false;
    if (n) {
      m_level = -1;
      if (skip_target)
        return false;
      m_level = 0;
      return true;
    }
    if (m_level < -1)
      return false;
  }
  m_level++;
  return true;
}

void ExitWithError(int code, LPCWSTR fmt, ...);

int TestContext::CheckEntry(const StackWalkerBase::TCallstackEntry & entry)
{
  if (m_level >= (int)m_callList.size())
    return 0;
  TestContext::CallData & cdata = m_callList[m_level];
  if (cdata.addr && entry.offset) {
    if ((LPVOID)entry.offset < cdata.addr && (LPBYTE)entry.offset > (LPBYTE)cdata.addr + 8)
      ExitWithError(1, L"Incorrect call addr in callstack. Expected: %p Received: %p \n", cdata.addr, (LPVOID)entry.offset);
  }
  if (entry.name == NULL || wcscmp(entry.name, cdata.name) != 0) {
    ExitWithError(1, L"Incorrect function name in callstack. Expected: \"%s\" \n", cdata.name);
  }
  if (cdata.line && entry.lineNumber) {
    if (entry.lineNumber != cdata.line && entry.lineNumber != cdata.line + 1)
      ExitWithError(1, L"Incorrect call line in callstack. Expected: %d or %d \n", cdata.line, cdata.line + 1);
  }
  return 1;
}

void TestContext::AddCall(LPVOID addr, LPCWSTR name, int line)
{
  m_callList.insert(m_callList.begin(), 1, CallData(addr, name, line));
}

// =========================================================================================

#define CALL(fn, ...)   ctx.AddCall((LPVOID)&fn,  __FUNCTIONW__, __LINE__); fn(__VA_ARGS__);
#define CALL2(fn, ...)  ctx.AddCall((LPVOID)NULL, __FUNCTIONW__, __LINE__); fn(__VA_ARGS__);

void InitTest(LPCSTR ns, LPCSTR caption)
{
  CHAR unit[MAX_PATH];
  strcpy_s(unit, __FILE__);
  LPSTR uname = strrchr(unit, '\\') ? strrchr(unit, '\\') + 1 : unit;
  printf("\n==========================================================\n");
  printf("Unit: %s, Run: '%s', Desc: \"%s\" \n", uname, ns, caption);
}

void CloseTest(LPCSTR ns, int level)
{
  if (level > 0 && level < 1000) {
    printf("[OK] Test \"%s\" finished. ++++++++++++++++++++++++++++++\n", ns);
  }
  if (level <= 0) {
    printf("[FAIL] Test \"%s\" ended incorrectly! \n", ns);
    ExitWithError(33, L"Level have incorrect value! (%d) \n", level);
  }
}

void ExitWithError(int code, LPCWSTR fmt, ...)
{
  va_list argptr;
  va_start(argptr, fmt);
  if (code > 10000)
    printf("FATAL ERROR: ");
  else if (code != 0)
    printf("ERROR: ");
  vwprintf_s(fmt, argptr);
  ExitProcess(code);
}

// =========================================================================================
namespace test1 {

const char caption[] = "Test RtlCaptureContext without exceptions.";

TestContext ctx;

int testCallstackEntry(const StackWalkerBase::TCallstackEntry & entry)
{
  if (!ctx.UpdateLevel(entry.name, L"::ShowCallstack"))
    return 0;
  return ctx.CheckEntry(entry);
}

void Func5()
{
  StackWalker sw;
  CALL2(sw.ShowCallstack, NULL, &ctx);
}

void Func4()
{  
  CALL(Func5);
}

void Func3()
{
  CALL(Func4);
}

void Func2()
{
  CALL(Func3);
}

void Func1()
{
  CALL(Func2);
}

int run()
{
  ctx.reset(-2, testCallstackEntry);
  CALL(Func1);
  return ctx.m_level;
}

} // namespace

// =========================================================================================
namespace test2 {

const char caption[] = "Test SEH context.";

TestContext ctx;

int testCallstackEntry(const StackWalkerBase::TCallstackEntry & entry)
{
  ctx.m_level++;
  return ctx.CheckEntry(entry);
}

LONG WINAPI ExpFilter(EXCEPTION_POINTERS * pExp, DWORD dwExpCode)
{
  StackWalker sw;
  sw.ShowCallstack(GetCurrentThread(), pExp->ContextRecord, NULL, &ctx);
  return EXCEPTION_EXECUTE_HANDLER;
}

__forceinline
void CreateException1()
{
  memset(NULL, 10, 10);
}

void Func4()
{  
  CALL2(CreateException1);
}

void Func3()
{
  CALL(Func4);
}

void Func2()
{
  CALL(Func3);
}

void Func1()
{
  CALL(Func2);
}

int run()
{
  ctx.reset(-1, testCallstackEntry);
  __try
  {
    CALL(Func1);
  }
  __except (ExpFilter(GetExceptionInformation(), GetExceptionCode()))
  {
    printf("Structured Exception Handler was called. \n");
  }
  return ctx.m_level;
}

} // namespace

// =========================================================================================
namespace test3 {

const char caption[] = "Test C++ exception context.";

TestContext ctx;

int testCallstackEntry(const StackWalkerBase::TCallstackEntry & entry)
{
  if (!ctx.UpdateLevel(entry.name, L"CxxThrowException"))
    return 0;
  return ctx.CheckEntry(entry);
}

__forceinline
void CreateException1()
{
  throw std::exception("fake exception"); 
}

void Func4()
{  
  CALL2(CreateException1);
}

void Func3()
{
  CALL(Func4);
}

void Func2()
{
  CALL(Func3);
}

void Func1()
{
  CALL(Func2);
}

int run()
{
  ctx.reset(-2, testCallstackEntry);
  try
  {
    CALL(Func1);
  }
  catch (...)
  {
    StackWalker sw;
    sw.ShowCallstack(sw.GetCurrentExceptionContext(), &ctx);
  }
  return ctx.m_level;
}

} // namespace

// =========================================================================================
namespace test4 {

const char caption[] = "Test reuse SW after loading new DLL.";

TestContext ctx;

int testCallstackEntry(const StackWalkerBase::TCallstackEntry & entry)
{
  if (!ctx.UpdateLevel(entry.name, L"DllGetVersion"))
    return 0;
  return ctx.CheckEntry(entry);
}

typedef VOID(WINAPI * FnDllGetVersion)(LPVOID pcdvi);
FnDllGetVersion DllGetVersion = NULL;

__forceinline
void CreateException1()
{
  if (DllGetVersion)
    DllGetVersion((LPVOID)8);
  else
    throw std::exception("fake exception");
}

void Func4()
{  
  CALL2(CreateException1);
}

void Func3()
{
  CALL(Func4);
}

void Func2()
{
  CALL(Func3);
}

void Func1()
{
  CALL(Func2);
}

int run()
{
  if (!g_EHasync) {
    printf("Skip test. Support only Async EH.\n");
    return 1000;
  }
  StackWalker sw;
  try
  {
    CALL(Func1);
  }
  catch (...)
  {
    sw.ShowCallstack(sw.GetCurrentExceptionContext());
  }
  HMODULE hLib = LoadLibraryA("cabinet.dll");
  DllGetVersion = (FnDllGetVersion)GetProcAddress(hLib, "DllGetVersion");
  ctx.reset(-1, testCallstackEntry);
  printf("===== call cabinet.DllGetVersion(bad_ptr) ======= \n");
  try
  {
    CALL(Func1);
  }
  catch (...)
  {
    sw.ShowCallstack(sw.GetCurrentExceptionContext(), &ctx);
  }
  return ctx.m_level;
}

} // namespace

// =========================================================================================

int CatchEHsync()
{
  try
  {
    memset(NULL, 10, 10);
  }
  catch (...)
  {
    g_EHasync = true;
    printf("############# Async exception handling detected [ /EHa ] #############\n");
  }
  return 0;
}

void DetectEHType()
{
  g_EHasync = false;
  __try
  {
    CatchEHsync();
  }
  __except(EXCEPTION_EXECUTE_HANDLER)
  {
    if (!g_EHasync)
      printf("############# Sync exception handling detected [ /EHs ] #############\n");
  }
}

// =========================================================================================

#define RUNTEST(ns, func, ...) \
  do { \
    InitTest(#ns, ns::caption); \
    int level = ns::func(__VA_ARGS__); \
    CloseTest(#ns, level); \
  } while(0)

int wmain(int argc, WCHAR * argv[])
{
  DetectEHType();
  RUNTEST(test1, run);
  RUNTEST(test2, run);
  RUNTEST(test3, run);
  RUNTEST(test4, run);
  return 0;
}

#endif // STKWLK_UNIT_TEST
