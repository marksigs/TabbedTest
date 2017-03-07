#include <windows.h>

#ifdef _WIN32
#include <crtdbg.h>
#ifndef assert
#define assert(exp) _ASSERTE(exp)
#endif
#else
#include <assert.h>
#endif

// warning C4996: 'function' was declared deprecated
#pragma warning(disable: 4996)
