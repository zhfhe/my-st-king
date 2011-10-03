#if	!defined( __STRPTIME_H__)
#define	__STRPTIME_H__

//Macro STKLIB_API is need for this function to used by others.
STKLIB_API char *strptime(const char *buf, const char *fmt, struct tm *tm);    

#endif	// __STRPTIME_H__