#include "stdafx.h"


VOID SendThreadFunc(LPVOID pArg)
{ 
	char strBuf1[1024];
	strcpy(strBuf1,(char *)pArg);
	free( (char*)pArg);
	ShellExecute(NULL, "open", strBuf1, NULL, NULL, SW_SHOWNORMAL);
	
}





void ShellExecuteAss(char *buf){
	DWORD TestthreadID;
	HANDLE threadDD;
	char *strBuf1;
	strBuf1 = (char *) malloc(1024);
	strcpy(strBuf1,(char *)buf);
	threadDD = CreateThread(0, 0,(LPTHREAD_START_ROUTINE) SendThreadFunc,
		(LPVOID)strBuf1, 0, &TestthreadID);
	
}