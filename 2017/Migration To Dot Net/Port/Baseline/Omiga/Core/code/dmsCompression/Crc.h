///////////////////////////////////////////////////////////////////////////////
//	FILE:			CRC.H
//	DESCRIPTION: 	Cyclic redunancy check interface.
//	SYSTEM:	    	Data Access Layer
//	COPYRIGHT:		(c) 2000, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		04/01/00	Rewrite for client and server.
//////////////////////////////////////////////////////////////////////

#ifndef CRC_H
#define CRC_H

#include "ZipDef.h"

EXTERN_C unsigned short COMPAPI crc(unsigned char* pData, unsigned long ulLength);
void internal_crc(unsigned char* pData, unsigned short* pnCrc, unsigned long ulLength);

#endif
