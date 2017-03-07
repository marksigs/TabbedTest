#ifdef EXPORT_ZIP
	#ifndef EXTERN_C
		#ifdef __cplusplus
			#define EXTERN_C extern "C"
		#else
			#define EXTERN_C extern
		#endif
	#endif
#else
	#ifdef EXTERN_C
		#undef EXTERN_C
	#endif
	#define	EXTERN_C
#endif

#ifdef EXPORT_ZIP
	#ifdef _WIN32
		#ifdef BUILD_COMP_DLL
			#define COMPAPI __declspec(dllexport)
		#else 
			#define COMPAPI __declspec(dllimport)
		#endif
	#else
		#ifdef BUILD_COMP_DLL
			#define COMPAPI far _export pascal
		#else
			#define COMPAPI far _import pascal
		#endif
	#endif
#else
	#ifdef COMPAPI
		#undef COMPAPI
	#endif
	#define COMPAPI
#endif
