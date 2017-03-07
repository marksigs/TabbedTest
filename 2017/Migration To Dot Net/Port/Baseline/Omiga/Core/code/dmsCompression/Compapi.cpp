///////////////////////////////////////////////////////////////////////////////
//	FILE:			Compapi.cpp
//	DESCRIPTION: 	This is the original third party compression code, with
//					changes to implement reading and writing Dac streams,
//					and to make it thread safe (removing all globals and 
//					statics).
//	SYSTEM:	    	Data Access Layer
//	COPYRIGHT:		(c) 2000, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		04/01/00	Split out of compapi.c.
///////////////////////////////////////////////////////////////////////////////

/*@H************************ < COMPRESS API    > ****************************
*   $@(#) compapi.c,v 4.3d 90/01/18 03:00:00 don Release ^                  *
*                                                                           *
*   compress : compapi.c  <current version of compress algorithm>           *
*                                                                           *
*   port by  : Donald J. Gloistein                                          *
*                                                                           *
*   Source, Documentation, Object Code:                                     *
*   released to Public Domain.  This code is based on code as documented    *
*   below  release notes.                                                 *
*                                                                           *
*---------------------------  Module Description  --------------------------*
*   Contains source code for modified Lempel-Ziv method (LZW) compression   *
*   and decompression.                                                      *
*                                                                           *
*   This code module can be maintained to keep current on releases on the   *
*   Unix system. The command shell and dos modules can remain the same.     *
*                                                                           *
*--------------------------- Implementation Notes --------------------------*
*                                                                           *
*   compiled with : compress.h compress.fns compress.c                      *
*   linked with   : compress.obj compusi.obj                                *
*                                                                           *
*   problems:                                                               *
*                                                                           *
*                                                                           *
*   CAUTION: Uses a number of defines for access and speed. If you change   *
*            anything, make sure about side effects.                        *
*                                                                           *
* Compression:                                                              *
* Algorithm:  use open addressing double hashing (no chaining) on the       *
* prefix code / next character combination.  We do a variant of Knuth's     *
* algorithm D (vol. 3, sec. 6.4) along with G. Knott's relatively-prime     *
* secondary probe.  Here, the modular division first probe is gives way     *
* to a faster exclusive-or manipulation.                                    *
* Also block compression with an adaptive reset was used in original code,  *
* whereby the code table is cleared when the compression ration decreases   *
* but after the table fills.  This was removed from this edition. The table *
* is re-sized at this point when it is filled , and a special CLEAR code is *
* generated for the decompressor. This results in some size difference from *
* straight version 4.0 joe Release. But it is fully compatible in both v4.0 *
* and v4.01                                                                 *
*                                                                           *
* Decompression:                                                            *
* This routine adapts to the codes in the file building the "string" table  *
* on-the-fly; requiring no table to be stored in the compressed file.  The  *
* tables used herein are shared with those of the compress() routine.       *
*                                                                           *
*     Initials ---- Name ---------------------------------                  *
*      DjG          Donald J. Gloistein, current port to MsDos 16 bit       *
*                   Plus many others, see rev.hst file for full list        *
*      LvR          Lyle V. Rains, many thanks for improved implementation  *
*                   of the compression and decompression routines.          *
*************************************************************************@H*/

#include "stdhdr.h"
#define MAIN
#include "CompAPIZipper.h"
#include "CompApi.h"
#include "comp.h" /* contains the rest of the include file declarations */


// FILE *in,*out;
// int offset;
// long int in_count ;		   /* length of input */
// long int bytes_out;		   /* length of compressed output */

// INTCODE prefxcode, nextfree;
// INTCODE highcode;
// INTCODE maxcode;
// HASH hashsize;
// int	bits;


/*
 * The following two parameter tables are the hash table sizes and
 * maximum code values for various code bit-lengths.  The requirements
 * are that Hashsize[n] must be a prime number and Maxcode[n] must be less
 * than Maxhash[n].  Table occupancy factor is (Maxcode - 256)/Maxhash.
 * Note:  I am using a lower Maxcode for 16-bit codes in order to
 * keep the hash table size less than 64k entries.
 */
CONST HASH hs[] = {
  0x13FF,       /* 12-bit codes, 75% occupancy */
  0x26C3,       /* 13-bit codes, 80% occupancy */
  0x4A1D,       /* 14-bit codes, 85% occupancy */
  0x8D0D,       /* 15-bit codes, 90% occupancy */
  0xFFD9        /* 16-bit codes, 94% occupancy, 6% of code values unused */
};
#define Hashsize(maxb) (hs[(maxb) -MINBITS])

CONST INTCODE mc[] = {
  0x0FFF,       /* 12-bit codes */
  0x1FFF,       /* 13-bit codes */
  0x3FFF,       /* 14-bit codes */
  0x7FFF,       /* 15-bit codes */
  0xEFFF        /* 16-bit codes, 6% of code values unused */
};
#define Maxcode(maxb) (mc[(maxb) -MINBITS])


ALLOCTYPE FAR *emalloc(unsigned int x, int y)
{
    ALLOCTYPE FAR *p;
    p = (ALLOCTYPE FAR *)ALLOCATE(x,y);
    return(p);
}

void efree(ALLOCTYPE FAR *ptr)
{
    FREEIT(ptr);
}


#ifdef __STDC__
#if 0 /*DEBUG*/
#define allocx(type, ptr, size) \
    (((ptr) = (type FAR *) emalloc((unsigned int)(size),sizeof(type))) == NULLPTR(type) \
    ?   (fprintf(stderr,"%s: "#ptr" -- ", prog_name), NOMEM) : OK \
    )
#else
#define allocx(type,ptr,size) \
    (((ptr) = (type FAR *) emalloc((unsigned int)(size),sizeof(type))) == NULLPTR(type) \
    ? NOMEM : OK \
    )
#endif
#else
#define allocx(type,ptr,size) \
    (((ptr) = (type FAR *) emalloc((unsigned int)(size),sizeof(type))) == NULLPTR(type) \
    ? NOMEM : OK \
    )
#endif

#define free_array(type,ptr,offset) \
    if (ptr != NULLPTR(type)) { \
        efree((ALLOCTYPE FAR *)((ptr) + (offset))); \
        (ptr) = NULLPTR(type); \
    }

  /*
   * Macro to allocate new memory to a pointer with an offset value.
   */
#define alloc_array(type, ptr, size, offset) \
    ( allocx(type, ptr, (size) - (offset)) != OK \
      ? NOMEM \
      : (((ptr) -= (offset)), OK) \
    )

//static char FAR *sfx = NULLPTR(char) ;
#define suffix(code)     pZipper->sfx[code]


#if (SPLIT_PFX)
	//static CODE FAR *pfx[2] = {NULLPTR(CODE), NULLPTR(CODE)};
#else
	//static CODE FAR *pfx = NULLPTR(CODE);
#endif

#if (SPLIT_HT)
	//static CODE FAR *ht[2] = {NULLPTR(CODE),NULLPTR(CODE)};
#else
	//static CODE FAR *ht = NULLPTR(CODE);
#endif


int alloc_tables(CCompAPIZipper* pZipper, INTCODE maxcode, HASH hashsize)
{
	//static INTCODE oldmaxcode = 0;
	//static HASH oldhashsize = 0;

	if (hashsize > pZipper->oldhashsize) 
	{
#if (SPLIT_HT)
		free_array(CODE,pZipper->ht[1], 0);
		free_array(CODE,pZipper->ht[0], 0);
#else
		free_array(CODE,pZipper->ht, 0);
#endif
		pZipper->oldhashsize = 0;
	}

    if (maxcode > pZipper->oldmaxcode) 
	{
#if (SPLIT_PFX)
        free_array(CODE,pZipper->pfx[1], 128);
        free_array(CODE,pZipper->pfx[0], 128);
#else
        free_array(CODE,pZipper->pfx, 256);
#endif
        free_array(char,pZipper->sfx, 256);

        if (   alloc_array(char, pZipper->sfx, maxcode + 1, 256)
#if (SPLIT_PFX)
            || alloc_array(CODE, pZipper->pfx[0], (maxcode + 1) / 2, 128)
            || alloc_array(CODE, pZipper->pfx[1], (maxcode + 1) / 2, 128)
#else
            || alloc_array(CODE, pZipper->pfx, (maxcode + 1), 256)
#endif
        ) 
		{
            pZipper->oldmaxcode = 0;
            pZipper->exit_stat = NOMEM;
            return(NOMEM);
        }
        pZipper->oldmaxcode = maxcode;
    }
    if (hashsize > pZipper->oldhashsize) 
	{
        if (
#if (SPLIT_HT)
            alloc_array(CODE, pZipper->ht[0], (hashsize / 2) + 1, 0)
            || alloc_array(CODE, pZipper->ht[1], hashsize / 2, 0)
#else
            alloc_array(CODE, pZipper->ht, hashsize, 0)
#endif
        ) {
            pZipper->oldhashsize = 0;
            pZipper->exit_stat = NOMEM;
            return(NOMEM);
        }
        pZipper->oldhashsize = hashsize;
    }
    return (OK);
}

# if (SPLIT_PFX)
    /*
     * We have to split pZipper->pfx[] table in half,
     * because it's potentially larger than 64k bytes.
     */
#   define prefix(code)   (pZipper->pfx[(code) & 1][(code) >> 1])
# else
    /*
     * Then pZipper->pfx[] can't be larger than 64k bytes,
     * or we don't care if it is, so we don't split.
     */
#   define prefix(code) (pZipper->pfx[code])
# endif


/* The initializing of the tables can be done quicker with memset() */
/* but this way is portable through out the memory models.          */
/* If you use Microsoft halloc() to allocate the arrays, then       */
/* include the pragma #pragma function(memset)  and make sure that  */
/* the length of the memory block is not greater than 64K.          */
/* This also means that you MUST compile in a model that makes the  */
/* default pointers to be far pointers (compact or large models).   */
/* See the file COMPUSI.DOS to modify function emalloc().           */

# if (SPLIT_HT)
    /*
     * We have to split pZipper->ht[] hash table in half,
     * because it's potentially larger than 64k bytes.
     */
#   define probe(hash)    (pZipper->ht[(hash) & 1][(hash) >> 1])
#   define init_tables() \
    { \
      hash = pZipper->hashsize >> 1; \
      pZipper->ht[0][hash] = 0; \
      while (hash--) pZipper->ht[0][hash] = pZipper->ht[1][hash] = 0; \
      pZipper->highcode = ~(~(INTCODE)0 << (pZipper->bits = INITBITS)); \
      pZipper->nextfree = (pZipper->block_compress ? FIRSTFREE : 256); \
    }

# else

    /*
     * Then pZipper->ht[] can't be larger than 64k bytes,
     * or we don't care if it is, so we don't split.
     */
#   define probe(hash) (pZipper->ht[hash])
#   define init_tables() \
    { \
      hash = pZipper->hashsize; \
      while (hash--) pZipper->ht[hash] = 0; \
      pZipper->highcode = ~(~(INTCODE)0 << (pZipper->bits = INITBITS)); \
      pZipper->nextfree = (pZipper->block_compress ? FIRSTFREE : 256); \
    }

# endif

#ifdef COMP40
/* table clear for block compress */
/* this is for adaptive reset present in version 4.0 joe release */
/* DjG, sets it up and returns TRUE to compress and FALSE to not compress */
int cl_block(CCompAPIZipper* pZipper)
{
    register long int rat;

    checkpoint = pZipper->in_count + CHECK_GAP;
#if 0 /*DEBUG*/
  if ( debug ) {
        fprintf ( stderr, "count: %ld, ratio: ", pZipper->in_count );
        prratio ( stderr, pZipper->in_count, pZipper->bytes_out );
    fprintf ( stderr, "\n");
  }
#endif

    if(pZipper->in_count > 0x007fffff) { /* shift will overflow */
        rat = pZipper->bytes_out >> 8;
        if(rat == 0)       /* Don't divide by zero */
            rat = 0x7fffffff;
        else
            rat = pZipper->in_count / rat;
    }
    else
        rat = (pZipper->in_count << 8) / pZipper->bytes_out;  /* 8 fractional bits */

    if ( rat > ratio ){
        ratio = rat;
        return FALSE;
    }
    else {
        ratio = 0;
#if 0 /*DEBUG*/
        if(debug)
        fprintf ( stderr, "clear\n" );
#endif
		return TRUE;	/* clear the table */
	}
	return FALSE;		/* don't clear the table */
}
#endif

/*
* compress stdin to stdout
*
*/
void compress(CCompAPIZipper* pZipper)
{
    int c,adjbits;
    register HASH hash;
    register INTCODE code;
    HASH hashf[256];
	
    pZipper->maxcode = Maxcode(pZipper->maxbits);
    pZipper->hashsize = Hashsize(pZipper->maxbits);
	
#ifdef COMP40
	/* Only needed for adaptive reset */
    checkpoint = CHECK_GAP;
    ratio = 0;
#endif
	
    adjbits = pZipper->maxbits -10;
    
	for (c = 256; --c >= 0; )
	{
        hashf[c] = ((( c &0x7) << 7) ^ c) << adjbits;
    }
	
    pZipper->exit_stat = OK;
    
	if (alloc_tables(pZipper, pZipper->maxcode, pZipper->hashsize))  /* pZipper->exit_stat already set */
        return;
    
	init_tables();
    
	/* if not zcat or filter */
#if 0
	//FIXME
	if(is_list && !zcat_flg) 
	{  
		/* Open output file */
		if (freopen(ofname, WRITE_FILE_TYPE, stdout) == NULL) 
		{
			pZipper->exit_stat = NOTOPENED;
			return;
		}
		if (!quiet)
			fprintf(stderr, "%s: ",ifname);
		setvbuf(stdout,zbuf,_IOFBF,ZBUFSIZE);
	}
#endif
	/*
	* Check the input stream for previously seen strings.  We keep
	* adding characters to the previously seen prefix string until we
	* get a character which forms a new (unseen) string.  We then send
	* the code for the previously seen prefix string, and add the new
	* string to our tables.  The check for previous strings is done by
	* hashing.  If the code for the hash value is unused, then we have
	* a new string.  If the code is used, we check to see if the prefix
	* and suffix values match the current input; if so, we have found
	* a previously seen string.  Otherwise, we have a hash collision,
	* and we try secondary hash probes until we either find the current
	* string, or we find an unused entry (which indicates a new string).
	*/
	pZipper->bytes_out = 0L;      /* no 3-byte header mojo */
	pZipper->in_count = 1L;
	pZipper->offset = 0;
	
	if ((c = pZipper->GetChar()) == EOF) 
	{
		pZipper->exit_stat = pZipper->GetLastError() ? READERR : OK;
		return;
	}
	
	pZipper->prefxcode = (INTCODE)c;
	
	while ((c = pZipper->GetChar()) != EOF) 
	{
		pZipper->in_count++;
		hash = pZipper->prefxcode ^ hashf[c];
		/* I need to check that my hash value is within range
		* because my 16-bit hash table is smaller than 64k.
		*/
		if (hash >= pZipper->hashsize)
			hash -= pZipper->hashsize;

		if ((code = (INTCODE)probe(hash)) != UNUSED) 
		{
			if (suffix(code) != (char)c || (INTCODE)prefix(code) != pZipper->prefxcode) 
			{
			/* hashdelta is subtracted from hash on each iteration of
			* the following hash table search loop.  I compute it once
			* here to remove it from the loop.
				*/
				HASH hashdelta = (0x120 - c) << (adjbits);
				do  {
					/* rehash and keep looking */
					assert(code >= FIRSTFREE && code <= pZipper->maxcode);
					if (hash >= hashdelta) hash -= hashdelta;
					else hash += (pZipper->hashsize - hashdelta);
					assert(hash < pZipper->hashsize);
					if ((code = (INTCODE)probe(hash)) == UNUSED)
						goto newcode;
				} while (suffix(code) != (char)c || (INTCODE)prefix(code) != pZipper->prefxcode);
			}
			pZipper->prefxcode = code;
		}
		else 
		{
newcode: 
			{
			putcode(pZipper, pZipper->prefxcode, pZipper->bits);
			code = pZipper->nextfree;
			assert(hash < pZipper->hashsize);
			assert(code >= FIRSTFREE);
			assert(code <= pZipper->maxcode + 1);
			if (code <= pZipper->maxcode) 
			{
				probe(hash) = (CODE)code;
				prefix(code) = (CODE)pZipper->prefxcode;
				suffix(code) = (char)c;
				if (code > pZipper->highcode) 
				{
					pZipper->highcode += code;
					++pZipper->bits;
				}
				pZipper->nextfree = code + 1;
			}
#ifdef COMP40
			else if (pZipper->in_count >= checkpoint && pZipper->block_compress ) 
			{
				if (cl_block(pZipper))
				{
#else
					else if (pZipper->block_compress)
					{
#endif
						putcode(pZipper, (INTCODE)c, pZipper->bits);
						putcode(pZipper, CLEAR, pZipper->bits);
						init_tables();

						if ((c = pZipper->GetChar()) == EOF)
							break;
						
						pZipper->in_count++;
#ifdef COMP40
					}
#endif
				}
				pZipper->prefxcode = (INTCODE)c;
			}
		}
	}

	putcode(pZipper, pZipper->prefxcode, pZipper->bits);
	putcode(pZipper, CLEAR, 0);

	if (pZipper->GetLastError())
	{ 
		/* check it on exit */
		pZipper->exit_stat = WRITEERR;
		return;
	}

	/*
	* Print out stats on stderr
	*/
#if 0
	if(zcat_flg == 0 && !quiet) 
	{
#if 0 /*DEBUG*/
		fprintf( stderr,
			"%ld chars in, (%ld bytes) out, compression factor: ",
			pZipper->in_count, pZipper->bytes_out );
		prratio( stderr, pZipper->in_count, pZipper->bytes_out );
		fprintf( stderr, "\n");
		fprintf( stderr, "\tCompression as in compact: " );
		prratio( stderr, pZipper->in_count-pZipper->bytes_out, pZipper->in_count );
		fprintf( stderr, "\n");
		fprintf( stderr, "\tLargest code (of last block) was %d (%d bits)\n",
			pZipper->prefxcode - 1, pZipper->bits );
#else
		fprintf( stderr, "Compression: " );
		prratio( stderr, pZipper->in_count-pZipper->bytes_out, pZipper->in_count );
#endif /* DEBUG */
	}
#endif

	if(pZipper->bytes_out > pZipper->in_count)      /*  if no savings */
		pZipper->exit_stat = NOSAVING;

	return ;
}
		

CONST UCHAR rmask[9] = {0x00, 0x01, 0x03, 0x07, 0x0f, 0x1f, 0x3f, 0x7f, 0xff};

void putcode(CCompAPIZipper* pZipper, INTCODE code, register int bits)
{
  //static int oldbits = 0;
  //static UCHAR outbuf[MAXBITS];
  register UCHAR *buf;
  register int shift;

  if (bits != pZipper->oldbits) {
    if (bits == 0) {
      /* bits == 0 means EOF, write the rest of the buffer. */
      if (pZipper->offset > 0) 
	  {
          pZipper->Write((char*)pZipper->outbuf, (pZipper->offset +7) >> 3);
          pZipper->bytes_out += ((pZipper->offset +7) >> 3);
      }
      pZipper->offset = 0;
      pZipper->oldbits = 0;
      return;
    }
    else {
      /* Change the code size.  We must write the whole buffer,
       * because the expand side won't discover the size change
       * until after it has read a buffer full.
       */
      if (pZipper->offset > 0) 
	  {
        pZipper->Write((char*)pZipper->outbuf, pZipper->oldbits);
        pZipper->bytes_out += pZipper->oldbits;
        pZipper->offset = 0;
      }
      pZipper->oldbits = bits;
#if 0 /*DEBUG*/
      if ( debug )
        fprintf( stderr, "\nChange to %d bits\n", bits );
#endif /* DEBUG */
    }
  }
  /*  Get to the first byte. */
  buf = pZipper->outbuf + ((shift = pZipper->offset) >> 3);
  if ((shift &= 7) != 0) {
    *(buf) |= (*buf & rmask[shift]) | (UCHAR)(code << shift);
    *(++buf) = (UCHAR)(code >> (8 - shift));
    if (bits + shift > 16)
        *(++buf) = (UCHAR)(code >> (16 - shift));
  }
  else {
    /* Special case for fast execution */
    *(buf) = (UCHAR)code;
    *(++buf) = (UCHAR)(code >> 8);
  }
  if ((pZipper->offset += bits) == (bits << 3)) {
    pZipper->bytes_out += bits;
    pZipper->Write((char*)pZipper->outbuf, bits);
    pZipper->offset = 0;
  }
  return;
}

///////////////////////////////////////////////////////////////////////////////
// Get the next code from input and put it in *codeptr.
// Return (TRUE) on success, or return (FALSE) on end-of-file.
// Adapted from COMPRESS V4.0.
///////////////////////////////////////////////////////////////////////////////
int nextcode(CCompAPIZipper* pZipper, INTCODE *codeptr)
{
	//static int prevbits = 0;
	register INTCODE code;
	//static int size;
	//static UCHAR inbuf[MAXBITS];
	register int shift;
	UCHAR *bp;
	
	/* If the next entry is a different bit-pZipper->size than the preceeding one
	* then we must adjust the pZipper->size and scrap the old buffer.
	*/
	if (pZipper->prevbits != pZipper->bits) 
	{
		pZipper->prevbits = pZipper->bits;
		pZipper->size = 0;
	}

	/* If we can't read another code from the buffer, then refill it.
	*/
	if (pZipper->size - (shift = pZipper->offset) < pZipper->bits) 
	{
		/* Read more input and convert pZipper->size from # of bytes to # of bits */
		if ((pZipper->size = pZipper->Read((char*)pZipper->inbuf, pZipper->bits) << 3) <= 0 || pZipper->GetLastError())
		{
			return(NO);
		}

		pZipper->offset = shift = 0;
	}
	
	/* Get to the first byte. */
	bp = pZipper->inbuf + (shift >> 3);
	
	/* Get first part (low order bits) */
	code = (*bp++ >> (shift &= 7));
	
	/* high order bits. */
	code |= *bp++ << (shift = 8 - shift);
	if ((shift += 8) < pZipper->bits) code |= *bp << shift;
	*codeptr = code & pZipper->highcode;
	pZipper->offset += pZipper->bits;
	
	return (TRUE);
}

void decompress(CCompAPIZipper* pZipper)
{
	register int i;
	register INTCODE code;
	char sufxchar;
	INTCODE savecode;
	FLAG fulltable, cleartable;
	//static char *token= NULL;         /* String buffer to build token */
	//static int maxtoklen = MAXTOKLEN;

	pZipper->exit_stat = OK;
	
	/* Initialze the pZipper->token buffer. */
    if (pZipper->token == NULL && (pZipper->token = (char*)malloc(pZipper->maxtoklen)) == NULL) 
	{
		pZipper->exit_stat = NOMEM;
		return;
	}
	
	if (alloc_tables(pZipper, pZipper->maxcode = ~(~(INTCODE)0 << pZipper->maxbits),0)) /* pZipper->exit_stat already set */
		return;
	
    /* if not zcat or filter */
#if 0
	//fixme
	if(is_list && !zcat_flg) 
	{  /* Open output file */
		if (freopen(ofname, WRITE_FILE_TYPE, stdout) == NULL) 
		{
			pZipper->exit_stat = NOTOPENED;
			return;
		}
		if (!quiet)
			fprintf(stderr, "%s: ",ifname);
		setvbuf(stdout,xbuf,_IOFBF,XBUFSIZE);
	}
#endif
	cleartable = TRUE;
	savecode = CLEAR;
	pZipper->offset = 0;
	do {
		if ((code = savecode) == CLEAR && cleartable) 
		{
			pZipper->highcode = ~(~(INTCODE)0 << (pZipper->bits = INITBITS));
			fulltable = FALSE;
			pZipper->nextfree = (cleartable = pZipper->block_compress) == FALSE ? 256 : FIRSTFREE;
			
			if (!nextcode(pZipper, &pZipper->prefxcode))
				break;

			sufxchar = (char)pZipper->prefxcode;
			pZipper->Write(&sufxchar, 1);
			continue;
		}
		i = 0;
		if (code >= pZipper->nextfree && !fulltable) {
			if (code != pZipper->nextfree){
				pZipper->exit_stat = CODEBAD;
				return ;     /* Non-existant code */
			}
			/* Special case for sequence KwKwK (see text of article)         */
			code = pZipper->prefxcode;
			pZipper->token[i++] = sufxchar;
		}

		/* Build the pZipper->token string in reverse order by chasing down through
		* successive prefix tokens of the current pZipper->token.  Then output it.
		*/
		while (code >= 256) {
#     if 0 /*DEBUG*/
		/* These are checks to ease paranoia. Prefix codes must decrease
		* monotonically, otherwise we must have corrupt tables.  We can
		* also check that we haven't overrun the pZipper->token buffer.
			*/
			if (code <= (INTCODE)prefix(code))
			{
				pZipper->exit_stat= TABLEBAD;
				return;
			}
#     endif
			if (i >= pZipper->maxtoklen) 
			{
				pZipper->maxtoklen *= 2;   /* double the size of the pZipper->token buffer */
				if ((pZipper->token = (char*)realloc(pZipper->token, pZipper->maxtoklen)) == NULL) 
				{
					pZipper->exit_stat = TOKTOOBIG;
					return;
				}
			}
			pZipper->token[i++] = suffix(code);
			code = (INTCODE)prefix(code);
		}

		sufxchar = (char)code;
		pZipper->Write(&sufxchar, 1);
		while (--i >= 0)
			pZipper->Write(&pZipper->token[i], 1);
			
		if (pZipper->GetLastError())
		{
			pZipper->exit_stat = WRITEERR;
			return;
		}

		if (pZipper->GetLastError()) 
		{
			pZipper->exit_stat = WRITEERR;
			return;
		}

		/* If table isn't full, add new pZipper->token code to the table with
		* codeprefix and codesuffix, and remember current code.
		*/
		if (!fulltable) 
		{
			code = pZipper->nextfree;
			assert(256 <= code && code <= pZipper->maxcode);
			prefix(code) = (CODE)pZipper->prefxcode;
			suffix(code) = sufxchar;
			pZipper->prefxcode = savecode;
			if (code++ == pZipper->highcode) 
			{
				if (pZipper->highcode >= pZipper->maxcode) 
				{
					fulltable = TRUE;
					--code;
				}
				else 
				{
					++pZipper->bits;
					pZipper->highcode += code;           /* pZipper->nextfree == pZipper->highcode + 1 */
				}
			}
			pZipper->nextfree = code;
		}
	} while (nextcode(pZipper, &savecode));

	pZipper->exit_stat = (pZipper->GetLastError())? READERR : OK;
	
	return ;
}
