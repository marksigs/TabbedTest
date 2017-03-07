#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <errno.h>
#include <fcntl.h>

#ifdef unix
# include <unistd.h>
# include <utime.h>
#else
# include <direct.h>
# include <io.h>
#endif

#include "unzip.h"

#ifdef ZLIBZIP
#define MINIUNZ
#include "zlibHelper.h"
#endif

#define CASESENSITIVITY (0)
#define WRITEBUFFERSIZE (8192)
#define MAXFILENAME (256)

#ifdef WIN32
#define USEWIN32IOAPI
#include "iowin32.h"
#endif
/*
  mini unzip, demo of unzip package

  usage :
  Usage : miniunz [-exvlo] file.zip [file_to_extract]

  list the file in the zipfile, and print the content of FILE_ID.ZIP or README.TXT
    if it exists
*/


/* change_file_date : change the date/time of a file
    filename : the filename of the file where date/time must be modified
    dosdate : the new date at the MSDos format (4 bytes)
    tmu_date : the SAME new date at the tm_unz format */
void change_file_date(filename,dosdate,tmu_date)
    const char *filename;
    uLong dosdate;
    tm_unz tmu_date;
{
#ifdef WIN32
  HANDLE hFile;
  FILETIME ftm,ftLocal,ftCreate,ftLastAcc,ftLastWrite;

  hFile = CreateFile(filename,GENERIC_READ | GENERIC_WRITE,
                      0,NULL,OPEN_EXISTING,0,NULL);
  GetFileTime(hFile,&ftCreate,&ftLastAcc,&ftLastWrite);
  DosDateTimeToFileTime((WORD)(dosdate>>16),(WORD)dosdate,&ftLocal);
  LocalFileTimeToFileTime(&ftLocal,&ftm);
  SetFileTime(hFile,&ftm,&ftLastAcc,&ftm);
  CloseHandle(hFile);
#else
#ifdef unix
  struct utimbuf ut;
  struct tm newdate;
  newdate.tm_sec = tmu_date.tm_sec;
  newdate.tm_min=tmu_date.tm_min;
  newdate.tm_hour=tmu_date.tm_hour;
  newdate.tm_mday=tmu_date.tm_mday;
  newdate.tm_mon=tmu_date.tm_mon;
  if (tmu_date.tm_year > 1900)
      newdate.tm_year=tmu_date.tm_year - 1900;
  else
      newdate.tm_year=tmu_date.tm_year ;
  newdate.tm_isdst=-1;

  ut.actime=ut.modtime=mktime(&newdate);
  utime(filename,&ut);
#endif
#endif
}


/* mymkdir and change_file_date are not 100 % portable
   As I don't know well Unix, I wait feedback for the unix portion */

int mymkdir(dirname)
    const char* dirname;
{
    int ret=0;
#ifdef WIN32
    ret = mkdir(dirname);
#else
#ifdef unix
    ret = mkdir (dirname,0775);
#endif
#endif
    return ret;
}

int makedir (newdir)
    char *newdir;
{
  char *buffer ;
  char *p;
  int  len = (int)strlen(newdir);

  if (len <= 0)
    return 0;

  buffer = (char*)malloc(len+1);
  strcpy(buffer,newdir);

  if (buffer[len-1] == '/') {
    buffer[len-1] = '\0';
  }
  if (mymkdir(buffer) == 0)
    {
      free(buffer);
      return 1;
    }

  p = buffer+1;
  while (1)
    {
      char hold;

      while(*p && *p != '\\' && *p != '/')
        p++;
      hold = *p;
      *p = 0;
      if ((mymkdir(buffer) == -1) && (errno == ENOENT))
        {
          printf("couldn't create directory %s\n",buffer);
          free(buffer);
          return 0;
        }
      if (hold == 0)
        break;
      *p++ = hold;
    }
  free(buffer);
  return 1;
}

void do_banner()
{
    printf("MiniUnz 1.00, demo of zLib + Unz package written by Gilles Vollant\n");
    printf("more info at http://www.winimage.com/zLibDll/unzip.html\n\n");
}

void do_help()
{
    printf("Usage : miniunz [-e] [-x] [-v] [-l] [-o] [-p password] [-d unzipDir] file.zip [file_to_extr.]\n\n" \
           "  -e  Extract without pathname (junk paths)\n" \
           "  -x  Extract with pathname\n" \
           "  -v  list files\n" \
           "  -l  list files\n" \
           "  -o  overwrite files without prompting\n" \
           "  -p  extract crypted file using password\n" \
		   "  -d  directory in which to unzip files\n\n");
}


int do_list(uf)
    unzFile uf;
{
    uLong i;
    unz_global_info gi;
    int err;

    err = unzGetGlobalInfo (uf,&gi);
    if (err!=UNZ_OK)
        printf("error %d with zipfile in unzGetGlobalInfo \n",err);
    printf(" Length  Method   Size  Ratio   Date    Time   CRC-32     Name\n");
    printf(" ------  ------   ----  -----   ----    ----   ------     ----\n");
    for (i=0;i<gi.number_entry;i++)
    {
        char filename_inzip[256];
        unz_file_info file_info;
        uLong ratio=0;
        const char *string_method;
        char charCrypt=' ';
        err = unzGetCurrentFileInfo(uf,&file_info,filename_inzip,sizeof(filename_inzip),NULL,0,NULL,0);
        if (err!=UNZ_OK)
        {
            printf("error %d with zipfile in unzGetCurrentFileInfo\n",err);
            break;
        }
        if (file_info.uncompressed_size>0)
            ratio = (file_info.compressed_size*100)/file_info.uncompressed_size;

        /* display a '*' if the file is crypted */
        if ((file_info.flag & 1) != 0)
            charCrypt='*';

        if (file_info.compression_method==0)
            string_method="Stored";
        else
        if (file_info.compression_method==Z_DEFLATED)
        {
            uInt iLevel=(uInt)((file_info.flag & 0x6)/2);
            if (iLevel==0)
              string_method="Defl:N";
            else if (iLevel==1)
              string_method="Defl:X";
            else if ((iLevel==2) || (iLevel==3))
              string_method="Defl:F"; /* 2:fast , 3 : extra fast*/
        }
        else
            string_method="Unkn. ";

        printf("%7lu  %6s%c%7lu %3lu%%  %2.2lu-%2.2lu-%2.2lu  %2.2lu:%2.2lu  %8.8lx   %s\n",
                file_info.uncompressed_size,string_method,
                charCrypt,
                file_info.compressed_size,
                ratio,
                (uLong)file_info.tmu_date.tm_mon + 1,
                (uLong)file_info.tmu_date.tm_mday,
                (uLong)file_info.tmu_date.tm_year % 100,
                (uLong)file_info.tmu_date.tm_hour,(uLong)file_info.tmu_date.tm_min,
                (uLong)file_info.crc,filename_inzip);
        if ((i+1)<gi.number_entry)
        {
            err = unzGoToNextFile(uf);
            if (err!=UNZ_OK)
            {
                printf("error %d with zipfile in unzGoToNextFile\n",err);
                break;
            }
        }
    }

    return 0;
}


int do_extract_currentfile(uf,popt_extract_without_path,popt_overwrite,password,unzipDir,pUnzippedFiles,plUnzippedFilesMaxSize,plUnzippedFilesCount)
    unzFile uf;
    const int* popt_extract_without_path;
    int* popt_overwrite;
    const char* password;
	const char* unzipDir;
	char* pUnzippedFiles;
	long* plUnzippedFilesMaxSize;
	long* plUnzippedFilesCount;
{
    char filename_inzip[256];
    char filename_with_zipdir[256];
    char* filename_withoutpath;
    char* p;
    int err=UNZ_OK;
	int errMem=UNZ_OK;
    FILE *fout=NULL;
    void* buf=NULL;
    uInt size_buf;
	char* pUnzippedFilesTemp = pUnzippedFiles;
	long lUnzippedFileSize = 0;
	long lUnzippedFilesMaxSize = 0;

    unz_file_info file_info;
    uLong ratio=0;
    err = unzGetCurrentFileInfo(uf,&file_info,filename_inzip,sizeof(filename_inzip),NULL,0,NULL,0);

    if (err!=UNZ_OK)
    {
        printf("error %d with zipfile in unzGetCurrentFileInfo\n",err);
        return err;
    }

    size_buf = WRITEBUFFERSIZE;
    buf = (void*)malloc(size_buf);
    if (buf==NULL)
    {
        printf("Error allocating memory\n");
        return UNZ_INTERNALERROR;
    }

	// AS 08/12/2005 CORE224
	// The unzipped files buffer (pUnzippedFiles) has the following structure for each unzipped file
	// it contains:
	//
	// Unzipped file name (null terminated string).
	// 4 bytes (long): Number of bytes in unzipped file.
	// Unzipped file (binary).
	//
	// The unzipped files occur immediately one after the other in the buffer.

	// Calculate the current size of the unzipped files buffer, 
	// and set pUnzippedFilesTemp to point to the end.
	if (plUnzippedFilesCount != NULL && pUnzippedFiles != NULL && plUnzippedFilesMaxSize != NULL)
	{
		int nUnzippedFile = 0;

		for (nUnzippedFile = 0; nUnzippedFile < *plUnzippedFilesCount; nUnzippedFile++)
		{
			long* pUnzippedFileSize = NULL;
			pUnzippedFilesTemp += strlen(pUnzippedFilesTemp) + 1;
			pUnzippedFileSize = (long*)pUnzippedFilesTemp;
			pUnzippedFilesTemp += sizeof(long) + (*pUnzippedFileSize);
		}

		lUnzippedFilesMaxSize = (long)((pUnzippedFilesTemp - pUnzippedFiles) + strlen(filename_inzip) + 1);
		if (lUnzippedFilesMaxSize > *plUnzippedFilesMaxSize)
		{
			errMem = UNZ_MOREMEM;
		}
		else
		{
			strcpy(pUnzippedFilesTemp, filename_inzip);
		}
		pUnzippedFilesTemp += strlen(filename_inzip) + 1;
	}

	p = filename_withoutpath = filename_inzip;
    while ((*p) != '\0')
    {
        if (((*p)=='/') || ((*p)=='\\'))
            filename_withoutpath = p+1;
        p++;
    }

    if ((*filename_withoutpath)=='\0')
    {
        if ((*popt_extract_without_path)==0)
        {
            printf("creating directory: %s\n",filename_inzip);
            mymkdir(filename_inzip);
        }
    }
    else
    {
        const char* write_filename;
        int skip=0;

        if ((*popt_extract_without_path)==0)
		{
			if (unzipDir != NULL)
			{
				strcpy(filename_with_zipdir, unzipDir);
				strcat(filename_with_zipdir, "\\");
				strcat(filename_with_zipdir, filename_withoutpath);
				write_filename = filename_with_zipdir;
			}
			else
			{
	            write_filename = filename_inzip;
			}
		}
		else
            write_filename = filename_withoutpath;

        err = unzOpenCurrentFilePassword(uf,password);
        if (err!=UNZ_OK)
        {
            printf("error %d with zipfile in unzOpenCurrentFilePassword\n",err);
        }

        if (((*popt_overwrite)==0) && (err==UNZ_OK) && pUnzippedFiles==NULL)
        {
            char rep=0;
            FILE* ftestexist;
            ftestexist = fopen(write_filename,"rb");
            if (ftestexist!=NULL)
            {
                fclose(ftestexist);
                do
                {
                    char answer[128];
                    printf("The file %s exist. Overwrite ? [y]es, [n]o, [A]ll: ",write_filename);
                    scanf("%1s",answer);
                    rep = answer[0] ;
                    if ((rep>='a') && (rep<='z'))
                        rep -= 0x20;
                }
                while ((rep!='Y') && (rep!='N') && (rep!='A'));
            }

            if (rep == 'N')
                skip = 1;

            if (rep == 'A')
                *popt_overwrite=1;
        }

        if ((skip==0) && (err==UNZ_OK) && pUnzippedFiles==NULL)
        {
            fout=fopen(write_filename,"wb");

            /* some zipfile don't contain directory alone before file */
            if ((fout==NULL) && ((*popt_extract_without_path)==0) &&
                                (filename_withoutpath!=(char*)filename_inzip))
            {
                char c=*(filename_withoutpath-1);
                *(filename_withoutpath-1)='\0';
                makedir(write_filename);
                *(filename_withoutpath-1)=c;
                fout=fopen(write_filename,"wb");
            }

            if (fout==NULL)
            {
                printf("error opening %s\n",write_filename);
            }
        }

        if (fout!=NULL || pUnzippedFiles!=NULL)
        {
			char* pCopied = NULL;

            printf(" extracting: %s\n",write_filename);
			
			// AS 08/12/2005 CORE224
			// Is there room for lUnzippedFileSize?
			if (pUnzippedFiles != NULL && plUnzippedFilesMaxSize != NULL)
			{
				lUnzippedFilesMaxSize = (long)((pUnzippedFilesTemp - pUnzippedFiles) + (long)sizeof(lUnzippedFileSize));
				if (lUnzippedFilesMaxSize > *plUnzippedFilesMaxSize)
				{
					errMem = UNZ_MOREMEM;
				}
				else
				{
					pCopied = pUnzippedFilesTemp;
					memcpy(pCopied, &lUnzippedFileSize, sizeof(lUnzippedFileSize));
				}
				pUnzippedFilesTemp += sizeof(lUnzippedFileSize);
			}

            do
            {
                err = unzReadCurrentFile(uf,buf,size_buf);
                if (err<0)
                {
                    printf("error %d with zipfile in unzReadCurrentFile\n",err);
                    break;
                }
                if (err>0)
				{
					if (pUnzippedFiles != NULL && plUnzippedFilesMaxSize != NULL)
					{
						// AS 08/12/2005 CORE224
						// Unzipping to memory, not file.
						lUnzippedFilesMaxSize = (long)((pUnzippedFilesTemp - pUnzippedFiles) + err);
						if (errMem == UNZ_OK)
						{
							if (lUnzippedFilesMaxSize > *plUnzippedFilesMaxSize)
							{
								// Buffer is not big enough, so do not copy into it.
								errMem = UNZ_MOREMEM;
							}
							else
							{
								memcpy(pUnzippedFilesTemp, buf, err);
							}
						}
						pUnzippedFilesTemp += err;
						lUnzippedFileSize += err;
					}
                    else if (fwrite(buf,err,1,fout)!=1)
                    {
                        printf("error in writing extracted file\n");
                        err=UNZ_ERRNO;
                        break;
                    }
				}
            }
            while (err>0);
            if (fout)
                    fclose(fout);

			if (pCopied != NULL)
			{
				memcpy(pCopied, &lUnzippedFileSize, sizeof(lUnzippedFileSize));
			}

            if (err==0 && pUnzippedFiles == NULL)
                change_file_date(write_filename,file_info.dosDate,
                                 file_info.tmu_date);
        }

        if (err==UNZ_OK)
        {
            err = unzCloseCurrentFile (uf);
            if (err!=UNZ_OK)
            {
                printf("error %d with zipfile in unzCloseCurrentFile\n",err);
            }
        }
        else
            unzCloseCurrentFile(uf); /* don't lose the error */
    }

    free(buf);

	if (err == UNZ_OK)
	{
		if (pUnzippedFiles != NULL && pUnzippedFilesTemp != NULL && plUnzippedFilesMaxSize != NULL)
		{
			// Set the actual size of buffer.
			*plUnzippedFilesMaxSize = (long)(pUnzippedFilesTemp - pUnzippedFiles);
		}

		if (plUnzippedFilesCount != NULL)
		{
			(*plUnzippedFilesCount)++;
		}
		err = errMem;
	}
	
    return err;
}


int do_extract(uf,opt_extract_without_path,opt_overwrite,password,unzipDir,pUnzippedFiles,plUnzippedFilesMaxSize,plUnzippedFilesCount)
    unzFile uf;
    int opt_extract_without_path;
    int opt_overwrite;
    const char* password;
	const char* unzipDir;
	char* pUnzippedFiles;
	long* plUnzippedFilesMaxSize;
	long* plUnzippedFilesCount;
{
    uLong i;
    unz_global_info gi;
    int err = UNZ_OK;
    FILE* fout=NULL;
	long lUnzippedFilesMaxSize = 0;

    err = unzGetGlobalInfo (uf,&gi);
    if (err!=UNZ_OK)
        printf("error %d with zipfile in unzGetGlobalInfo \n",err);

    for (i=0;i<gi.number_entry;i++)
    {
		// AS 08/12/2005 CORE224
		// Only interested in the final value;
		lUnzippedFilesMaxSize = plUnzippedFilesMaxSize ? *plUnzippedFilesMaxSize : 0;

		err = do_extract_currentfile(uf,&opt_extract_without_path,
                                      &opt_overwrite,
                                      password,
									  unzipDir,
									  pUnzippedFiles,
									  &lUnzippedFilesMaxSize,
									  plUnzippedFilesCount);

		// AS 08/12/2005 CORE224
		// If err == UNZ_MOREMEM then keep on unzipping the remaining files, as we need to find out
		// exactly how big the buffer needs to be, and return this to the caller in lUnzippedFilesMaxSize.
        if (err != UNZ_OK && err != UNZ_MOREMEM)
            break;

        if ((i+1)<gi.number_entry)
        {
            int errGoto = unzGoToNextFile(uf);
            if (errGoto!=UNZ_OK)
            {
                printf("error %d with zipfile in unzGoToNextFile\n",err);
				err = errGoto;
                break;
            }
        }
    }

	if (plUnzippedFilesMaxSize != NULL)
	{
		// AS 08/12/2005 CORE224
		// Return to the caller the maximum size of the pUnzippedFiles buffer.
		// If the buffer passed in was bigger, then the buffer will contain all the unzipped files, and
		// plUnzippedFilesMaxSize will be how many bytes in the buffer were actually used.
		// If the buffer passed in was too small, then the buffer will contain only the unzipped files that
		// could fit in the buffer, and plUnzippedFilesMaxSize will specify how big the buffer needs to be
		// to contain all the unzipped files.
		*plUnzippedFilesMaxSize = lUnzippedFilesMaxSize;
	}

    return err;
}

int do_extract_onefile(uf,filename,opt_extract_without_path,opt_overwrite,password,unzipDir,pUnzippedFiles,plUnzippedFilesMaxSize,plUnzippedFilesCount)
    unzFile uf;
    const char* filename;
    int opt_extract_without_path;
    int opt_overwrite;
    const char* password;
	const char* unzipDir;
	char* pUnzippedFiles;
	long* plUnzippedFilesMaxSize;
	long* plUnzippedFilesCount;
{
    int err = UNZ_OK;
    if (unzLocateFile(uf,filename,CASESENSITIVITY)!=UNZ_OK)
    {
        printf("file %s not found in the zipfile\n",filename);
        return 2;
    }

    if (do_extract_currentfile(uf,&opt_extract_without_path,
                                      &opt_overwrite,
                                      password,
									  unzipDir,
									  pUnzippedFiles,
									  plUnzippedFilesMaxSize,
									  plUnzippedFilesCount) == UNZ_OK)
        return 0;
    else
        return 1;
}


int mainEx(argc,argv,pUnzippedFiles,plUnzippedFilesMaxSize,plUnzippedFilesCount)
    int argc;
    char *argv[];
	char* pUnzippedFiles;
	long* plUnzippedFilesMaxSize;
	long* plUnzippedFilesCount;
{
    const char *zipfilename=NULL;
    const char *filename_to_extract=NULL;
    const char *password=NULL;
	const char *unzipDir=NULL;
    char filename_try[MAXFILENAME+16] = "";
    int i;
    int opt_do_list=0;
    int opt_do_extract=1;
    int opt_do_extract_withoutpath=0;
    int opt_overwrite=0;
	int result=0;
	int errclose;
    unzFile uf=NULL;

    do_banner();
    if (argc==1)
    {
        do_help();
        return 0;
    }
    else
    {
        for (i=1;i<argc;i++)
        {
            if ((*argv[i])=='-')
            {
                const char *p=argv[i]+1;

                while ((*p)!='\0')
                {
                    char c=*(p++);;
                    if ((c=='l') || (c=='L'))
                        opt_do_list = 1;
                    if ((c=='v') || (c=='V'))
                        opt_do_list = 1;
                    if ((c=='x') || (c=='X'))
                        opt_do_extract = 1;
                    if ((c=='e') || (c=='E'))
                        opt_do_extract = opt_do_extract_withoutpath = 1;
                    if ((c=='o') || (c=='O'))
                        opt_overwrite=1;
                    if (((c=='p') || (c=='P')) && (i+1<argc))
                    {
                        password=argv[i+1];
                        i++;
                    }
                    if (((c=='d') || (c=='D')) && (i+1<argc))
                    {
                        unzipDir=argv[i+1];
                        i++;
                    }
                }
            }
            else
            {
                if (zipfilename == NULL)
                    zipfilename = argv[i];
                else if (filename_to_extract==NULL)
                        filename_to_extract = argv[i] ;
            }
        }
    }

    if (zipfilename!=NULL)
    {

#        ifdef USEWIN32IOAPI
        zlib_filefunc_def ffunc;
#        endif

        strncpy(filename_try, zipfilename,MAXFILENAME-1);
        /* strncpy doesnt append the trailing NULL, of the string is too long. */
        filename_try[ MAXFILENAME ] = '\0';

#        ifdef USEWIN32IOAPI
        fill_win32_filefunc(&ffunc);
        uf = unzOpen2(zipfilename,&ffunc);
#        else
        uf = unzOpen(zipfilename);
#        endif
        if (uf==NULL)
        {
            strcat(filename_try,".zip");
#            ifdef USEWIN32IOAPI
            uf = unzOpen2(filename_try,&ffunc);
#            else
            uf = unzOpen(filename_try);
#            endif
        }
    }

    if (uf==NULL)
    {
        printf("Cannot open %s or %s.zip\n",zipfilename,zipfilename);
        return 1;
    }
    printf("%s opened\n",filename_try);

    if (opt_do_list==1)
        result = do_list(uf);
    else if (opt_do_extract==1)
    {
        if (filename_to_extract == NULL)
            result = do_extract(uf,opt_do_extract_withoutpath,opt_overwrite,password,unzipDir,pUnzippedFiles,plUnzippedFilesMaxSize,plUnzippedFilesCount);
        else
            result = do_extract_onefile(uf,filename_to_extract,
                                      opt_do_extract_withoutpath,opt_overwrite,password,unzipDir,pUnzippedFiles,plUnzippedFilesMaxSize,plUnzippedFilesCount);
    }

    errclose = unzClose(uf);
    if (errclose != UNZ_OK)
        printf("error in closing %s\n",filename_try);

    return result;
}

int main(argc,argv)
    int argc;
    char *argv[];
{
	return mainEx(argc,argv,NULL,NULL,NULL);
}
