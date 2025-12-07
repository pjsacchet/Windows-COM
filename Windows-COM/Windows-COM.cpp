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
    OLECHAR* writeStr = (OLECHAR*)L"Write";
    DISPID dispid;    
    VARIANT result = { VT_EMPTY }, filenameVar = { VT_EMPTY }, overwriteVar = { VT_EMPTY }, unicodeVar = { VT_EMPTY }, textVar = { VT_EMPTY };
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

    printf("Successfully created instance of FileSystem interface object\n");

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

    printf("Successfully created test file\n");

    // Now should have a TextStream object returned to us, so we should be able to write to it 

    // We get back a dispatch interface for our TextStream object, so we should be able to GetIDsOfNames again and Invoke our Write method with it 
    status = result.pdispVal->GetIDsOfNames(IID_NULL, &writeStr, 1, LOCALE_USER_DEFAULT, &dispid);
    if (status != S_OK)
    {
        printf("Failed GetIDsofNames; error 0x%X\n", status);
        goto cleanup;
    }

    printf("DISPID for Write is %d\n", dispid);

    textVar.vt = VT_BSTR;
    textVar.bstrVal = SysAllocString(L"This is a test string!");

    // Reposition our arg
    args[0] = textVar;
    args[1] = { 0 };
    args[2] = { 0 };

    params.cArgs = 1;

    status = result.pdispVal->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, &exceptInfo, &argErr);
    if (status != S_OK)
    {
        printf("Failed Invoke; error 0x%X\n", status);
        goto cleanup;
    }

    printf("Successfully wrote to output file\n");


cleanup:
    return status;
}


int main()
{
    HRESULT status = S_OK;
    
    printf("Loading COM library...\n");

    // Initialize COM library
    status = CoInitialize(NULL);
    if (status != S_OK)
    {
        printf("Failed CoInitialize; error 0x%X\n", status);
        return -1;
    }

    printf("Initializing file system object...\n");

    // Create our file
    CreateFileCOM();

    printf("Cleaning up...\n");

    // Unload COM library
    CoUninitialize();

    
    return ERROR_SUCCESS;
}
