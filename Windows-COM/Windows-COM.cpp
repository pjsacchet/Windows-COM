// Patrick Sacchet
// Baseline entry file for COM playground

#include <iostream>
#include <comdef.h>

#include "IFile.h"
#include "IFileSystem.h"


// Create a file in a particular directory via IFileSystem->IDispatch interface
    // This interface is exposed by Windows' Visual Basic for Applications (VBA) libraries
HRESULT CreateFileCOM()
{
    HRESULT status = S_OK;
    IDispatch *fileSystemDispatchInterface = (IDispatch*) NULL;
    OLECHAR *createTextFileStr = (OLECHAR*)L"CreateTextFile";
    DISPID dispid;    
    VARIANT result = { VT_EMPTY }, filenameVar = { VT_EMPTY }, overwriteVar = { VT_EMPTY }, unicodeVar = { VT_EMPTY };
    DISPPARAMS params = { 0 };
    EXCEPINFO exceptInfo = { 0 };
    UINT argErr = 0;
    VARIANTARG args[3];

    // First get a handle to IDspatch -> can create via IUknown -> QueryInterface or just create an instance here
    status = CoCreateInstance(fileSystemCLSID, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (PVOID*) &fileSystemDispatchInterface);
    if (status != S_OK)
    {
        printf("Failed CoCreateInstance; error 0x%X\n", status);
        goto cleanup;
    }

    // Call GetIDsOfNames to get the dispatch ID (DISPID) for IFileSystem->CreateTextFile()
    status = fileSystemDispatchInterface->GetIDsOfNames(IID_NULL, &createTextFileStr, 1, LOCALE_USER_DEFAULT, &dispid);
    if (status != S_OK)
    {
        printf("Failed GetIDsofNames; error 0x%X\n", status);
        goto cleanup;
    }

    printf("DISPID for CreateTextFile is %d\n", dispid);

    // Before we call CreateTextFile, we need to format our parameters correctly
    // Protocol for this method (outlined by OleView.NET) is as follows:
        // TextStream CreateTextFile(string FileName, [Optional] bool Overwrite, [Optional] bool Unicode);

    // Unicode bool
    unicodeVar.vt = VT_BOOL;
    unicodeVar.boolVal = FALSE;    

    // Overwrite bool
    overwriteVar.vt = VT_BOOL;
    overwriteVar.boolVal = TRUE;

    // Filename string
    filenameVar.vt = VT_BSTR;
    filenameVar.bstrVal = SysAllocString( L"C:\\test\\test.txt");

    // Important to note these arguments are stored in reverse order
    args[0] = unicodeVar;
    args[1] = overwriteVar;
    args[2] = filenameVar;
    
    params.cArgs = 3;
    params.rgvarg = args;

    // Now invoke our method with the required params
    status = fileSystemDispatchInterface->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, &exceptInfo, &argErr);
    if (status != S_OK)
    {
        printf("Failed Invoke; error 0x%X\n", status);
        goto cleanup;
    }

    // Now should have a TextStream object returned to us, so we should be able to write to it 
    

cleanup:
    return status;
}


int main()
{
    HRESULT status = S_OK;
    

    printf("Initializing file system...\n");

    // Initialize COM library
    status = CoInitialize(NULL);
    if (status != S_OK)
    {
        printf("Failed CoInitialize; error 0x%X\n", status);
        return -1;
    }

    // Create our file
    CreateFileCOM();


    // Unload COM library
    CoUninitialize();

    
    return ERROR_SUCCESS;
}
