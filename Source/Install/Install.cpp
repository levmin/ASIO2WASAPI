/*  ASIO2WASAPI Universal ASIO Driver
    Copyright (C) Lev Minkovsky
    
    This file is part of ASIO2WASAPI.

    ASIO2WASAPI is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    ASIO2WASAPI is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with ASIO2WASAPI; if not, write to the Free Software
    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
*/

#include "stdafx.h"
#include "..\Driver\version.h"

using namespace std;

//The install is a 32-bit executable. This function returns true if we run on 64-bit Windows
BOOL IsWow64()
{
    typedef BOOL (WINAPI *LPFN_ISWOW64PROCESS) (HANDLE, PBOOL);

    LPFN_ISWOW64PROCESS fnIsWow64Process;

    BOOL bIsWow64 = FALSE;

    fnIsWow64Process = (LPFN_ISWOW64PROCESS) GetProcAddress(
        GetModuleHandle(TEXT("kernel32")),"IsWow64Process");

    if(NULL != fnIsWow64Process)
    {
        fnIsWow64Process(GetCurrentProcess(),&bIsWow64);
    }
    return bIsWow64;
}

void AddTrailingSeparator(string& str)
{
   if (str.length()==0 || str[str.length()-1]!='\\')
      str+="\\";
}

int CopyFiles(string installFolder, const char * * names, int nNumberOfNames)
{
   AddTrailingSeparator(installFolder);

   for (int i=0;i<nNumberOfNames;i++)
   {
      DWORD attr=0;
      string fileName = installFolder+names[i];
      if (!CopyFile(names[i],fileName.c_str(),FALSE)
          || (attr=GetFileAttributes(fileName.c_str())==INVALID_FILE_ATTRIBUTES)
          || !SetFileAttributes(fileName.c_str(),attr & ~FILE_ATTRIBUTE_READONLY))
	      {
            MessageBox(NULL,(string("Copy operation failed for ")+names[i]).c_str(),NULL,MB_OK);
		      return -1;
	      }
      
   }
   return 0;
}

int RegisterDLLs(string installFolder, const char ** names, int nNumberOfDlls)
{
   AddTrailingSeparator(installFolder);
   for (int index=0;index<nNumberOfDlls;index++)
   {
      string dllName=installFolder+names[index];
      string commandLine = " /s \"" + dllName +"\"";
      ShellExecute(NULL,NULL,"regsvr32.exe",commandLine.c_str(),NULL,SW_HIDE);
   }
   return 0;
}

//set current directory to the location of the install executable
void SetFileDirectory()
{
   char image_dir[MAX_PATH] = {0};
   GetModuleFileName(NULL,image_dir,MAX_PATH);
   string imageDir(image_dir);
   int nPos = imageDir.find_last_of('\\');
   if (nPos < 0)
      return; //should not happen
   imageDir = imageDir.substr(0,nPos+1);
   SetCurrentDirectory(imageDir.c_str());
}

void GetEnvVar(const char * varName, string& var)
{
    const int BUF_LEN=32767; 
    char buf[BUF_LEN]={0}; 
    GetEnvironmentVariable(varName,buf,BUF_LEN);
    var = buf;
}

int WINAPI WinMain(HINSTANCE,HINSTANCE,LPSTR,int)
{
	if (!IsWindowsVistaOrGreater())
   {
      MessageBox(NULL,"ASIO2WASAPI required Windows Vista or later","ASIO2WASAPI Installer",MB_OK);
      return -1;
   }
   //set current directory to be the one of install.exe
   //this is needed because CopyFiles assumes the files are in the current directory
   SetFileDirectory();
   string programFiles;
   GetEnvVar("ProgramFiles",programFiles);
   AddTrailingSeparator(programFiles);
   string installFolder=programFiles+"ASIO2WASAPI";

   CreateDirectory(installFolder.c_str(),NULL);
   
   const char * fileNames[]={
      "ASIO2WASAPI.dll",
      "UnInstall.exe",
   };
   
   const char * dllNames[]={
      "ASIO2WASAPI.dll",
   };
   
   if (CopyFiles(installFolder,fileNames,sizeof(fileNames)/sizeof(fileNames[0]))!=0 ||
      RegisterDLLs(installFolder,dllNames,sizeof(dllNames)/sizeof(dllNames[0]))!=0)
      return -1;

   if (IsWow64())
   {
      SetCurrentDirectory("x64");
      string programFiles64;
      GetEnvVar("ProgramW6432",programFiles64);
      AddTrailingSeparator(programFiles64);
      string installFolder64=programFiles64+"ASIO2WASAPI";

      CreateDirectory(installFolder64.c_str(),NULL);
      const char * names[]={
         "ASIO2WASAPI64.dll",
      };
      if (CopyFiles(installFolder64,names,sizeof(names)/sizeof(names[0]))!=0 ||
         RegisterDLLs(installFolder64,names,sizeof(names)/sizeof(names[0])))
         return -1;
   }

   HKEY key = NULL;
   RegCreateKeyEx(HKEY_LOCAL_MACHINE,"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ASIO2WASAPI",0,0,0,KEY_ALL_ACCESS,0,&key,NULL);
   string uninstallString=installFolder;
   AddTrailingSeparator(uninstallString);
   uninstallString+="Uninstall.exe";
   struct {
      const char * pValueName;
      const char * pValueData;
   } values[] = 
      { 
         {"DisplayName","ASIO2WASAPI"},
         {"DisplayVersion",SPRODUCT_VERSION},
         {"Publisher","Lev Minkovsky"},
         {"HelpLink","http://asio2wasapi.sourceforge.net"},
         {"InstallLocation",installFolder.c_str()},
         {"Comments","Universal ASIO driver"},
         {"UninstallString",uninstallString.c_str()},
      };
   for (int i=0; i<sizeof(values)/sizeof(values[0]); i++)
      RegSetValueEx(key,values[i].pValueName,0,REG_SZ,(const BYTE *)values[i].pValueData,strlen(values[i].pValueData)+1);
   MessageBox(NULL,"ASIO2WASAPI is successfully installed","ASIO2WASAPI Installer",MB_OK);
   return 0;
}


