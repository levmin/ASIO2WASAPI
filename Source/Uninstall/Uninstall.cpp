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

using namespace std;

void AddTrailingSeparator(string& str)
{
   if (str.length()==0 || str[str.length()-1]!='\\')
      str+="\\";
}

//The uninstall is a 32-bit executable. This function returns true if we run on 64-bit Windows
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

int UnregisterDLLs(string installFolder, const char ** names, int nNumberOfDlls)
{
   AddTrailingSeparator(installFolder);

   for (int index=0;index<nNumberOfDlls;index++)
   {
      string dllName=installFolder+names[index];
      string commandLine = " /u /s \"" + dllName + "\"";
      ShellExecute(NULL,NULL,"regsvr32.exe",commandLine.c_str(),NULL,SW_HIDE);
   }
   return 0;
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
   string programFiles;
   GetEnvVar("ProgramFiles",programFiles);
   AddTrailingSeparator(programFiles);
   string installFolder=programFiles+"ASIO2WASAPI";

   string programFiles64;
   GetEnvVar("ProgramW6432",programFiles64);
   AddTrailingSeparator(programFiles64);
   string installFolder64=programFiles64+"ASIO2WASAPI";

   const char * names[]={
      "ASIO2WASAPI.dll",
   };
   if (UnregisterDLLs(installFolder,names,sizeof(names)/sizeof(names[0])))
      return -1;
   if (IsWow64())
   {
      const char * names[]={
         "ASIO2WASAPI64.dll",
      };
      if (UnregisterDLLs(installFolder64,names,sizeof(names)/sizeof(names[0])))
         return -1;
   }
   
   const char * uninstKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ASIO2WASAPI";
   RegDeleteKey(HKEY_LOCAL_MACHINE,uninstKey);

   //create and execute a bat file deleting install folders 
   char szTmpPath[MAX_PATH];
   GetTempPath(MAX_PATH,szTmpPath);
   string tmpDir(szTmpPath);
   AddTrailingSeparator(tmpDir);
   string tmpBatFile=tmpDir+"finish_cleanup.bat";
	string uninstallExe=installFolder+"\\Uninstall.exe";
	FILE * f=NULL;
   fopen_s(&f,tmpBatFile.c_str(),"wt");
	fputs("echo off\n",f);
	fputs(":wait\n",f);
	fputs(("del \""+uninstallExe+"\"\n").c_str(),f);
	fputs(("if exist \""+uninstallExe+"\" goto wait\n").c_str(),f);
	fputs(("rd /s /q \""+installFolder+"\"\n").c_str(),f);
	if (IsWow64())
		fputs(("rd /s /q \""+installFolder64+"\"\n").c_str(),f);
    fputs(("del \""+tmpDir+"finish_cleanup.bat\"\n").c_str(),f);
	fclose(f);
	ShellExecute(NULL, NULL, tmpBatFile.c_str(), NULL, NULL, SW_HIDE);

	return 0;
}
