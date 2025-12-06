// Patrick Sacchet
// Baseline entry file for COM playground

#include <iostream>

#include "IFile.h"
#include "IFileSystem.h"




int main()
{
    HRESULT status;
    PVOID fileSystemInterface;

    printf("Initializing file system...\n");

    status = CoInitialize(NULL);
    if (status != S_OK)
    {
        printf("Failed CoInitialize; error 0x%X\n", status);
        return -1;
    }

    // First need IFileSystem before we can create a file 
    status = CoCreateInstance(fileSystemCLSID, NULL, CLSCTX_INPROC_SERVER, fileSystemIID, &fileSystemInterface);
    if (status != S_OK)
    {
        printf("Failed CoCreateInstance; error 0x%X\n", status);
        return -1;
    }
    
    return ERROR_SUCCESS;
}
