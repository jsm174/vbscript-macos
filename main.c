#include "wine/debug.h"
#include "wine/test.h"

#include "vbscript.h"

#define IActiveScriptParse_QueryInterface IActiveScriptParse64_QueryInterface
#define IActiveScriptParse_Release IActiveScriptParse64_Release
#define IActiveScriptParse_InitNew IActiveScriptParse64_InitNew
#define IActiveScriptParse_ParseScriptText IActiveScriptParse64_ParseScriptText
#define IActiveScriptParseProcedure2_Release IActiveScriptParseProcedure2_64_Release

static BSTR a2bstr(const char *str)
{
    BSTR ret;
    int len;

    len = MultiByteToWideChar(CP_UTF8, 0, str, -1, NULL, 0);
    ret = SysAllocStringLen(NULL, len-1);
    MultiByteToWideChar(CP_UTF8, 0, str, -1, ret, len);

    return ret;
}

static const char *vt2a(VARIANT *v)
{
    if(V_VT(v) == (VT_BYREF|VT_VARIANT)) {
        static char buf[64];
        sprintf(buf, "%s*", vt2a(V_BYREF(v)));
        return buf;
    }

    switch(V_VT(v)) {
    case VT_EMPTY:
        return "VT_EMPTY";
    case VT_NULL:
        return "VT_NULL";
    case VT_I2:
        return "VT_I2";
    case VT_I4:
        return "VT_I4";
    case VT_R4:
        return "VT_R4";
    case VT_R8:
        return "VT_R8";
    case VT_CY:
        return "VT_CY";
    case VT_DATE:
        return "VT_DATE";
    case VT_BSTR:
        return "VT_BSTR";
    case VT_DISPATCH:
        return "VT_DISPATCH";
    case VT_UNKNOWN:
        return "VT_UNKNOWN";
    case VT_BOOL:
        return "VT_BOOL";
    case VT_ARRAY|VT_VARIANT:
        return "VT_ARRAY|VT_VARIANT";
    case VT_ARRAY|VT_BYREF|VT_VARIANT:
        return "VT_ARRAY|VT_BYREF|VT_VARIANT";
    case VT_UI1:
        return "VT_UI1";
    default:
        ok(0, "unknown vt %d\n", V_VT(v));
        return NULL;
    }
}


static HRESULT WINAPI ActiveScriptSite_QueryInterface(IActiveScriptSite *iface, REFIID riid, void **ppv)
{
    *ppv = NULL;

    if(IsEqualGUID(&IID_IUnknown, riid))
        *ppv = iface;
    else if(IsEqualGUID(&IID_IActiveScriptSite, riid))
        *ppv = iface;
    else
        return E_NOINTERFACE;

    IUnknown_AddRef((IUnknown*)*ppv);
    return S_OK;
}

static ULONG WINAPI ActiveScriptSite_AddRef(IActiveScriptSite *iface)
{
    return 2;
}

static ULONG WINAPI ActiveScriptSite_Release(IActiveScriptSite *iface)
{
    return 1;
}

static HRESULT WINAPI ActiveScriptSite_GetLCID(IActiveScriptSite *iface, LCID *plcid)
{
    return E_NOTIMPL;
}

static HRESULT WINAPI ActiveScriptSite_GetItemInfo(IActiveScriptSite *iface, LPCOLESTR pstrName,
        DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti)
{
    return S_OK;
}

static HRESULT WINAPI ActiveScriptSite_GetDocVersionString(IActiveScriptSite *iface, BSTR *pbstrVersion)
{
    return E_NOTIMPL;
}

static HRESULT WINAPI ActiveScriptSite_OnScriptTerminate(IActiveScriptSite *iface,
        const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo)
{
    return E_NOTIMPL;
}

static HRESULT WINAPI ActiveScriptSite_OnStateChange(IActiveScriptSite *iface, SCRIPTSTATE ssScriptState)
{
    return E_NOTIMPL;
}

static ULONG error_line;
static LONG error_char;

static HRESULT WINAPI ActiveScriptSite_OnScriptError(IActiveScriptSite *iface, IActiveScriptError *pscripterror)
{
    HRESULT hres;

    hres = IActiveScriptError_GetSourcePosition(pscripterror, NULL, &error_line, &error_char);
    
    EXCEPINFO info;

    hres = IActiveScriptError_GetExceptionInfo(pscripterror, &info);
        if(SUCCEEDED(hres))
            printf("Error in line %lu: %x %s\n", error_line + 1, info.wCode, wine_dbgstr_w(info.bstrDescription));


    ok(hres == S_OK, "GetSourcePosition failed: %08lx\n", hres);

    return S_OK;
}

static HRESULT WINAPI ActiveScriptSite_OnEnterScript(IActiveScriptSite *iface)
{
    return S_OK;
}

static HRESULT WINAPI ActiveScriptSite_OnLeaveScript(IActiveScriptSite *iface)
{
    return S_OK;
}

static const IActiveScriptSiteVtbl ActiveScriptSiteVtbl = {
    ActiveScriptSite_QueryInterface,
    ActiveScriptSite_AddRef,
    ActiveScriptSite_Release,
    ActiveScriptSite_GetLCID,
    ActiveScriptSite_GetItemInfo,
    ActiveScriptSite_GetDocVersionString,
    ActiveScriptSite_OnScriptTerminate,
    ActiveScriptSite_OnStateChange,
    ActiveScriptSite_OnScriptError,
    ActiveScriptSite_OnEnterScript,
    ActiveScriptSite_OnLeaveScript
};

static IActiveScriptSite ActiveScriptSite = { &ActiveScriptSiteVtbl };

int main(int argc, char **argv) { 
    IActiveScript *engine;
    IActiveScriptParse *parser;
    HRESULT hres;

    VBScriptFactory_CreateInstance(NULL, NULL, &IID_IActiveScript, (void**)&engine);

    if(!engine)
        return 0;

    hres = IActiveScript_QueryInterface(engine, &IID_IActiveScriptParse, (void**)&parser);
    hres = IActiveScriptParse_InitNew(parser);
    ok(hres == S_OK, "InitNew failed: %08lx\n", hres);

    IActiveScriptParse_Release(parser);

    hres = IActiveScript_SetScriptSite(engine, &ActiveScriptSite);
    ok(hres == S_OK, "SetScriptSite failed: %08lx\n", hres);
  
    hres = IActiveScript_AddNamedItem(engine, L"test",
         SCRIPTITEM_ISVISIBLE|SCRIPTITEM_ISSOURCE|0);
    ok(hres == S_OK, "AddNamedItem failed: %08lx\n", hres);

    hres = IActiveScript_SetScriptState(engine, SCRIPTSTATE_STARTED);
    ok(hres == S_OK, "SetScriptState(SCRIPTSTATE_STARTED) failed: %08lx\n", hres);

    VARIANT var;
    BSTR str;

    str = a2bstr("Function test(x)\n"
                 "    test = x + 0.5\n"
                 "End Function\n");
    hres = IActiveScriptParse_ParseScriptText(parser, str, NULL, NULL, NULL, 0, 0, 0, NULL, NULL);
    ok(hres == S_OK, "ParseScriptText failed: %08lx\n", hres);
    SysFreeString(str);

    str = a2bstr("test(4) * 3\n");
    hres = IActiveScriptParse_ParseScriptText(parser, str, NULL, NULL, NULL, 0, 0, SCRIPTTEXT_ISEXPRESSION, &var, NULL);
    ok(hres == S_OK, "ParseScriptText failed: %08lx\n", hres);
    ok(V_VT(&var) == VT_R8, "Expected VT_R8, got %s\n", vt2a(&var));
    ok(V_R8(&var) == 13.5, "Expected %lf, got %lf\n", 13.5, V_R8(&var));
    VariantClear(&var);
    SysFreeString(str);

    V_VT(&var) = VT_I2;
    str = a2bstr("If True Then foo = 42 Else foo = 0\n");
    hres = IActiveScriptParse_ParseScriptText(parser, str, NULL, NULL, NULL, 0, 0, 0, &var, NULL);
    ok(hres == S_OK, "ParseScriptText failed: %08lx\n", hres);
    ok(V_VT(&var) == VT_EMPTY, "Expected VT_EMPTY, got %s\n", vt2a(&var));
    VariantClear(&var);
    SysFreeString(str);

    str = a2bstr("foo\n\n");
    hres = IActiveScriptParse_ParseScriptText(parser, str, NULL, NULL, NULL, 0, 0, SCRIPTTEXT_ISEXPRESSION, &var, NULL);
    ok(hres == S_OK, "ParseScriptText failed: %08lx\n", hres);
    ok(V_VT(&var) == VT_I2, "Expected VT_I2, got %s\n", vt2a(&var));
    ok(V_I2(&var) == 42, "Expected 42, got %d\n", V_I2(&var));
    VariantClear(&var);
    SysFreeString(str);
}