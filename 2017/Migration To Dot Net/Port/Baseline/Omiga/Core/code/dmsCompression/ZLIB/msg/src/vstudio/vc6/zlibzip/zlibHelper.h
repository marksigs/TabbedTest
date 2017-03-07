#ifndef ZLIBHELPER_H
#define ZLIBHELPER_H

#include <stdio.h>


#ifdef __cplusplus
extern "C"
{
#endif
int printfZlibHelper(const char * format, ...);
#ifdef __cplusplus
};
#endif

#ifdef ZLIBZIP
// In the zlibzip project, replace printf (suitable for console application) with printfZlibHelper
// (suitable for a Win32 DLL).
#define printf printfZlibHelper
#endif

#if defined(MINIZIP)
#define do_banner do_banner_MINIZIP
#define do_help do_help_MINIZIP
#define main main_MINIZIP
#elif defined(MINIUNZ)
#define do_banner do_banner_MINIUNZ
#define do_help do_help_MINIUNZ
#define main main_MINIUNZ
#define mainEx main_MINIUNZEX
#endif

#endif
