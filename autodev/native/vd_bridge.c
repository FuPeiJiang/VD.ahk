/* VD for AutoDev: narrow x64 Windows COM adapter.
 * The user-facing library is VD.mjs. Windows' private desktop interfaces are
 * vtable-only (not IDispatch), so AutoDev's COM automation cannot invoke them.
 * This native adapter exposes only fixed, checked entry points for ffi.bind.
 * No long-lived COM references, callbacks, hooks, or global service pointers.
 */
#define WIN32_LEAN_AND_MEAN
#define _WIN32_WINNT 0x0A00
#include <windows.h>
#include <objbase.h>
#include <stdint.h>
#include <stdlib.h>
#include <wchar.h>
#include <winver.h>

#define EXPORT __declspec(dllexport)
#define SLOT(obj,n) (((void**)*(void**)(obj))[n])
#define RELEASE(p) do { if (p) { ((ULONG (STDMETHODCALLTYPE*)(void*))SLOT(p,2))(p); (p)=NULL; } } while(0)
#define ERR_UNSUPPORTED ((HRESULT)0x80004001L)
#define ERR_NOT_FOUND ((HRESULT)0x80070490L)

static HRESULT parse_guid(const wchar_t *s, GUID *out) { return CLSIDFromString(s, out); }
static const wchar_t *const clsid_shell=L"{c2f03a33-21f5-47fa-b4bb-156362a2f239}";
static const wchar_t *const iid_provider=L"{6d5140c1-7436-11ce-8034-00aa006009fa}";
static const wchar_t *const sid_mgr=L"{c5e0cdca-7b6e-41b2-9fc4-d93975cc467b}";
static const wchar_t *const sid_view=L"{1841c6d7-4f9d-42c0-af41-8747538f10e5}";
static const wchar_t *const sid_pinned=L"{b5a399e7-1c87-46b8-88e9-fc5747b171bd}";
static const wchar_t *const iid_pinned=L"{4ce81583-1e4c-4632-a621-07a53543148f}";
static const GUID guid_pin_app={0xbb64d5b7,0x4de3,0x4ab2,{0xa8,0x7c,0xdb,0x76,0x01,0xae,0xa7,0xdc}};
static const GUID guid_pin_view={0xc2ddea68,0x66f2,0x4cf9,{0x82,0x64,0x1b,0xfd,0x00,0xfb,0xbb,0xac}};

typedef struct Abi {
    uint32_t threshold_build,threshold_revision;
    const wchar_t *mgr_iid, *desktop_iid;
    int monitor;
    int move, current, desktops, sw, create, remove;
    int get_name, set_name, get_wall, set_wall;
} Abi;
/* ABI selection matches the reference VD.ahk: first threshold strictly greater
 * than twinui.pcshell.dll's build/revision; unsupported future builds fail closed. */
static const Abi abis[]={
 {20348,0,L"{f31574d6-b682-4cdc-bd56-1827860abec6}",L"{ff72ffdd-be7e-43fc-9c03-ad81681e88e4}",0,4,6,7,9,10,11,-1,-1,-1,-1},
 {22000,0,L"{094afe11-44f2-4ba0-976f-29a97e263ee0}",L"{62fdf88b-11ca-4afb-8bd8-2296dfae49e2}",1,4,6,7,9,10,11,6,14,-1,-1},
 {22483,0,L"{b2f925b9-5a0f-4d2e-9f4d-2b1507593c10}",L"{536d3495-b208-4cc9-ae26-de8111275bf8}",1,4,6,7,9,10,12,6,15,7,16},
 {22621,2215,L"{b2f925b9-5a0f-4d2e-9f4d-2b1507593c10}",L"{536d3495-b208-4cc9-ae26-de8111275bf8}",1,4,6,8,10,11,13,6,16,7,17},
 {22631,3085,L"{a3175f2d-239c-4bd2-8aa0-eeba8b0b138e}",L"{3f07f4be-b107-441a-af0f-39d82529072c}",0,4,6,7,9,10,12,5,15,6,16},
 {26100,0,L"{53f5ca0b-158f-4124-900c-057158060b27}",L"{3f07f4be-b107-441a-af0f-39d82529072c}",0,4,6,7,9,10,12,5,15,6,16},
 {99999,0,L"{53f5ca0b-158f-4124-900c-057158060b27}",L"{3f07f4be-b107-441a-af0f-39d82529072c}",0,4,6,7,9,11,13,5,16,6,17}
};

typedef struct Context {
    const Abi *abi;
    void *provider,*mgr,*views,*pinned;
    int initialized;
} Context;
static HRESULT system_version(uint32_t *build,uint32_t *revision) {
    wchar_t sys[MAX_PATH+64]; DWORD unused=0; UINT len=0; void *ptr=NULL;
    UINT n=GetSystemDirectoryW(sys,MAX_PATH);
    if (!n || n+20>=MAX_PATH+64) return E_FAIL;
    wcscat(sys,L"\\twinui.pcshell.dll");
    DWORD sz=GetFileVersionInfoSizeW(sys,&unused);
    if (!sz) return HRESULT_FROM_WIN32(GetLastError());
    void *data=malloc(sz);
    if (!data) return E_OUTOFMEMORY;
    HRESULT hr=E_FAIL;
    if (GetFileVersionInfoW(sys,0,sz,data) && VerQueryValueW(data,L"\\",&ptr,&len) && len>=sizeof(VS_FIXEDFILEINFO)) {
        const VS_FIXEDFILEINFO *v=(const VS_FIXEDFILEINFO*)ptr;
        *build=HIWORD(v->dwFileVersionLS); *revision=LOWORD(v->dwFileVersionLS);
        hr=S_OK;
    }
    free(data);return hr;
}
static HRESULT select_abi(const Abi **out) {
    uint32_t build=0,rev=0; HRESULT hr=system_version(&build,&rev);
    if (FAILED(hr)) return hr;
    for (unsigned i=0;i<sizeof(abis)/sizeof(abis[0]);i++) {
        if (build<abis[i].threshold_build || (build==abis[i].threshold_build && rev<abis[i].threshold_revision)) {
            *out=&abis[i]; return S_OK;
        }
    }
    return ERR_UNSUPPORTED;
}
static HRESULT service(void *provider,const wchar_t *sid,const wchar_t *iid,void **out) {
    GUID s,i;HRESULT hr=parse_guid(sid,&s); if(FAILED(hr))return hr;
    hr=parse_guid(iid,&i); if(FAILED(hr))return hr;
    return ((HRESULT (STDMETHODCALLTYPE*)(void*,const GUID*,const GUID*,void**))SLOT(provider,3))(provider,&s,&i,out);
}
static void end(Context *c) {
    RELEASE(c->pinned);RELEASE(c->views);RELEASE(c->mgr);RELEASE(c->provider);
    if(c->initialized)CoUninitialize();
}
static HRESULT begin(Context *c,int need_views,int need_pin) {
    ZeroMemory(c,sizeof(*c)); HRESULT hr=CoInitializeEx(NULL,COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(hr)) c->initialized=1;
    else if (hr!=RPC_E_CHANGED_MODE) return hr;
    hr=select_abi(&c->abi);if(FAILED(hr))return hr;
    GUID cls,iid;parse_guid(clsid_shell,&cls);parse_guid(iid_provider,&iid);
    hr=CoCreateInstance(&cls,NULL,CLSCTX_ALL,&iid,&c->provider);if(FAILED(hr))return hr;
    hr=service(c->provider,sid_mgr,c->abi->mgr_iid,&c->mgr);if(FAILED(hr))return hr;
    if(need_views) {hr=service(c->provider,sid_view,sid_view,&c->views);if(FAILED(hr))return hr;}
    if(need_pin) {hr=service(c->provider,sid_pinned,iid_pinned,&c->pinned);if(FAILED(hr))return hr;}
    return S_OK;
}
static HRESULT mgr_current(Context *c,void **out) {
    if(c->abi->monitor) return ((HRESULT (STDMETHODCALLTYPE*)(void*,void*,void**))SLOT(c->mgr,c->abi->current))(c->mgr,NULL,out);
    return ((HRESULT (STDMETHODCALLTYPE*)(void*,void**))SLOT(c->mgr,c->abi->current))(c->mgr,out);
}
static HRESULT mgr_desktops(Context *c,void **out) {
    if(c->abi->monitor) return ((HRESULT (STDMETHODCALLTYPE*)(void*,void*,void**))SLOT(c->mgr,c->abi->desktops))(c->mgr,NULL,out);
    return ((HRESULT (STDMETHODCALLTYPE*)(void*,void**))SLOT(c->mgr,c->abi->desktops))(c->mgr,out);
}
static HRESULT mgr_switch(Context *c,void *desk) {
    if(c->abi->monitor) return ((HRESULT (STDMETHODCALLTYPE*)(void*,void*,void*))SLOT(c->mgr,c->abi->sw))(c->mgr,NULL,desk);
    return ((HRESULT (STDMETHODCALLTYPE*)(void*,void*))SLOT(c->mgr,c->abi->sw))(c->mgr,desk);
}
static HRESULT mgr_create(Context *c,void **out) {
    if(c->abi->monitor) return ((HRESULT (STDMETHODCALLTYPE*)(void*,void*,void**))SLOT(c->mgr,c->abi->create))(c->mgr,NULL,out);
    return ((HRESULT (STDMETHODCALLTYPE*)(void*,void**))SLOT(c->mgr,c->abi->create))(c->mgr,out);
}
static HRESULT desks(Context *c,void **array,UINT *count) {
    HRESULT hr=mgr_desktops(c,array);if(FAILED(hr))return hr;
    return ((HRESULT (STDMETHODCALLTYPE*)(void*,UINT*))SLOT(*array,3))(*array,count);
}
static HRESULT desk_at(Context *c,void *array,UINT index,void **out) {
    GUID iid;parse_guid(c->abi->desktop_iid,&iid);
    return ((HRESULT (STDMETHODCALLTYPE*)(void*,UINT,const GUID*,void**))SLOT(array,4))(array,index,&iid,out);
}
static HRESULT view_from_hwnd(Context *c,uintptr_t hwnd,void **out) {
    if(!hwnd || !IsWindow((HWND)hwnd))return E_INVALIDARG;
    return ((HRESULT (STDMETHODCALLTYPE*)(void*,HWND,void**))SLOT(c->views,6))(c->views,(HWND)hwnd,out);
}
static HRESULT indexed(Context *c,int number,void **array,void **desktop,UINT *total) {
    *array=NULL;*desktop=NULL;*total=0;
    HRESULT hr=desks(c,array,total);if(FAILED(hr))return hr;
    if(number<1 || (UINT)number>*total)return E_INVALIDARG;
    return desk_at(c,*array,(UINT)(number-1),desktop);
}

EXPORT int32_t vd_version(int32_t *build,int32_t *revision,int32_t *abi) {
    const Abi *v=NULL;uint32_t b=0,r=0;
    if(!build||!revision||!abi)return E_POINTER;
    HRESULT hr=system_version(&b,&r);if(FAILED(hr))return hr;
    hr=select_abi(&v);if(FAILED(hr))return hr;
    *build=(int32_t)b;*revision=(int32_t)r;*abi=(int32_t)(v-abis);
    return S_OK;
}
EXPORT int32_t vd_count(int32_t *out) {
    if(!out)return E_POINTER;
    Context c;HRESULT hr=begin(&c,0,0);void *arr=NULL;UINT count=0;
    if(SUCCEEDED(hr))hr=desks(&c,&arr,&count);
    if(SUCCEEDED(hr))*out=(int32_t)count;
    RELEASE(arr);end(&c);return hr;
}
EXPORT int32_t vd_current(int32_t *out) {
    if(!out)return E_POINTER;
    Context c;HRESULT hr=begin(&c,0,0);void *arr=NULL,*now=NULL,*at=NULL;UINT count=0;
    if(SUCCEEDED(hr))hr=mgr_current(&c,&now);
    if(SUCCEEDED(hr))hr=desks(&c,&arr,&count);
    if(SUCCEEDED(hr)) {
        hr=ERR_NOT_FOUND;
        for(UINT i=0;i<count;i++) {
            if(SUCCEEDED(desk_at(&c,arr,i,&at))) {
                if(at==now) {*out=(int32_t)i+1;hr=S_OK;RELEASE(at);break;}
                RELEASE(at);
            }
        }
    }
    RELEASE(now);RELEASE(arr);end(&c);return hr;
}
EXPORT int32_t vd_switch(int32_t target) {
    Context c;HRESULT hr=begin(&c,0,0);void *arr=NULL,*desktop=NULL;UINT count=0;
    if(SUCCEEDED(hr))hr=indexed(&c,target,&arr,&desktop,&count);
    if(SUCCEEDED(hr))hr=mgr_switch(&c,desktop);
    RELEASE(desktop);RELEASE(arr);end(&c);return hr;
}
EXPORT int32_t vd_create(int32_t *count_after) {
    if(!count_after)return E_POINTER;
    Context c;HRESULT hr=begin(&c,0,0);void *created=NULL,*arr=NULL;UINT count=0;
    if(SUCCEEDED(hr))hr=mgr_create(&c,&created);
    if(SUCCEEDED(hr))hr=desks(&c,&arr,&count);
    if(SUCCEEDED(hr))*count_after=(int32_t)count;
    RELEASE(arr);RELEASE(created);end(&c);return hr;
}
EXPORT int32_t vd_remove(int32_t number,int32_t fallback) {
    Context c;HRESULT hr=begin(&c,0,0);void *arr=NULL,*d=NULL,*f=NULL;UINT count=0;
    if(SUCCEEDED(hr))hr=desks(&c,&arr,&count);
    if(SUCCEEDED(hr)) {
        if(count<2 || number<1 || (UINT)number>count)hr=E_INVALIDARG;
        else {
            if(fallback==0)fallback=(number==1)?2:number-1;
            if(fallback<1||(UINT)fallback>count||number==fallback)hr=E_INVALIDARG;
        }
    }
    if(SUCCEEDED(hr))hr=desk_at(&c,arr,(UINT)number-1,&d);
    if(SUCCEEDED(hr))hr=desk_at(&c,arr,(UINT)fallback-1,&f);
    if(SUCCEEDED(hr))hr=((HRESULT (STDMETHODCALLTYPE*)(void*,void*,void*))SLOT(c.mgr,c.abi->remove))(c.mgr,d,f);
    RELEASE(f);RELEASE(d);RELEASE(arr);end(&c);return hr;
}
EXPORT int32_t vd_move_window(uintptr_t hwnd,int32_t target) {
    Context c;HRESULT hr=begin(&c,1,0);void *arr=NULL,*d=NULL,*view=NULL;UINT count=0;
    if(SUCCEEDED(hr))hr=indexed(&c,target,&arr,&d,&count);
    if(SUCCEEDED(hr))hr=view_from_hwnd(&c,hwnd,&view);
    if(SUCCEEDED(hr))hr=((HRESULT (STDMETHODCALLTYPE*)(void*,void*,void*))SLOT(c.mgr,c.abi->move))(c.mgr,view,d);
    RELEASE(view);RELEASE(d);RELEASE(arr);end(&c);return hr;
}
EXPORT int32_t vd_window_desktop(uintptr_t hwnd,int32_t *out) {
    if(!out)return E_POINTER;
    Context c;HRESULT hr=begin(&c,1,0);void *view=NULL,*arr=NULL,*d=NULL;UINT count=0;GUID id;
    if(SUCCEEDED(hr))hr=view_from_hwnd(&c,hwnd,&view);
    if(SUCCEEDED(hr))hr=((HRESULT (STDMETHODCALLTYPE*)(void*,GUID*))SLOT(view,25))(view,&id);
    if(SUCCEEDED(hr)) {
        if(IsEqualGUID(&id,&guid_pin_app)) {*out=-1;goto done_window;}
        if(IsEqualGUID(&id,&guid_pin_view)) {*out=-2;goto done_window;}
        hr=desks(&c,&arr,&count);
        if(SUCCEEDED(hr)) {
            hr=ERR_NOT_FOUND;
            for(UINT i=0;i<count;i++) {
                if(SUCCEEDED(desk_at(&c,arr,i,&d))) {
                    GUID did;
                    /* IVirtualDesktop::GetID is slot 4, following IsViewVisible (slot 3). */
                    HRESULT h2=((HRESULT (STDMETHODCALLTYPE*)(void*,GUID*))SLOT(d,4))(d,&did);
                    RELEASE(d);
                    if(SUCCEEDED(h2)&&IsEqualGUID(&did,&id)){*out=(int32_t)i+1;hr=S_OK;break;}
                }
            }
        }
    }
done_window:
    RELEASE(d);RELEASE(arr);RELEASE(view);end(&c);return hr;
}
EXPORT int32_t vd_is_pinned(uintptr_t hwnd,int32_t *out) {
    if(!out)return E_POINTER;
    Context c;HRESULT hr=begin(&c,1,1);void *v=NULL;BOOL pinned=FALSE;
    if(SUCCEEDED(hr))hr=view_from_hwnd(&c,hwnd,&v);
    if(SUCCEEDED(hr))hr=((HRESULT (STDMETHODCALLTYPE*)(void*,void*,BOOL*))SLOT(c.pinned,6))(c.pinned,v,&pinned);
    if(SUCCEEDED(hr))*out=pinned?1:0;
    RELEASE(v);end(&c);return hr;
}
EXPORT int32_t vd_pin_window(uintptr_t hwnd,int32_t pin) {
    Context c;HRESULT hr=begin(&c,1,1);void *v=NULL;
    if(SUCCEEDED(hr))hr=view_from_hwnd(&c,hwnd,&v);
    if(SUCCEEDED(hr))hr=((HRESULT (STDMETHODCALLTYPE*)(void*,void*))SLOT(c.pinned,pin?7:8))(c.pinned,v);
    RELEASE(v);end(&c);return hr;
}
EXPORT int32_t vd_is_app_pinned(const wchar_t *app_id,int32_t *out) {
    if(!app_id||!out)return E_POINTER;
    Context c;HRESULT hr=begin(&c,0,1);BOOL pinned=FALSE;
    if(SUCCEEDED(hr))hr=((HRESULT (STDMETHODCALLTYPE*)(void*,const wchar_t*,BOOL*))SLOT(c.pinned,3))(c.pinned,app_id,&pinned);
    if(SUCCEEDED(hr))*out=pinned?1:0;
    end(&c);return hr;
}
EXPORT int32_t vd_pin_app(const wchar_t *app_id,int32_t pin) {
    if(!app_id)return E_POINTER;
    Context c;HRESULT hr=begin(&c,0,1);
    if(SUCCEEDED(hr))hr=((HRESULT (STDMETHODCALLTYPE*)(void*,const wchar_t*))SLOT(c.pinned,pin?4:5))(c.pinned,app_id);
    end(&c);return hr;
}
/* kind 0 = desktop name; kind 1 = wallpaper path. UTF-16 buffer, chars incl NUL. */
EXPORT int32_t vd_get_text(int32_t number,int32_t kind,wchar_t *dst,int32_t capacity) {
    if(!dst||capacity<=0||(kind!=0&&kind!=1))return E_INVALIDARG;
    dst[0]=0;
    Context c;HRESULT hr=begin(&c,0,0);void *arr=NULL,*d=NULL,*hstr=NULL;UINT count=0;
    if(SUCCEEDED(hr))hr=indexed(&c,number,&arr,&d,&count);
    if(SUCCEEDED(hr)) {
        int slot=kind?c.abi->get_wall:c.abi->get_name;
        if(slot<0)hr=ERR_UNSUPPORTED;
        else hr=((HRESULT (STDMETHODCALLTYPE*)(void*,void**))SLOT(d,slot))(d,&hstr);
    }
    if(SUCCEEDED(hr)) {
        HMODULE combase=GetModuleHandleW(L"combase.dll");
        if(!combase)combase=LoadLibraryW(L"combase.dll");
        typedef const wchar_t *(WINAPI *GetStringFn)(void*,UINT*);
        GetStringFn get=(GetStringFn)GetProcAddress(combase,"WindowsGetStringRawBuffer");
        if(!get)hr=ERR_UNSUPPORTED;
        else {
            UINT len=0;const wchar_t *txt=get(hstr,&len);
            if(len>=(UINT)capacity)hr=HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
            else {if(len&&txt)memcpy(dst,txt,len*sizeof(wchar_t));dst[len]=0;}
        }
        typedef HRESULT (WINAPI *DeleteStringFn)(void*);
        DeleteStringFn del=(DeleteStringFn)GetProcAddress(combase,"WindowsDeleteString");
        if(del&&hstr)del(hstr);
    }
    RELEASE(d);RELEASE(arr);end(&c);return hr;
}
EXPORT int32_t vd_set_text(int32_t number,int32_t kind,const wchar_t *src) {
    if(!src||(kind!=0&&kind!=1)||wcslen(src)>2048)return E_INVALIDARG;
    Context c;HRESULT hr=begin(&c,0,0);void *arr=NULL,*d=NULL,*hstr=NULL;UINT count=0;
    if(SUCCEEDED(hr))hr=indexed(&c,number,&arr,&d,&count);
    if(SUCCEEDED(hr)) {
        int slot=kind?c.abi->set_wall:c.abi->set_name;
        if(slot<0)hr=ERR_UNSUPPORTED;
        else {
            HMODULE combase=GetModuleHandleW(L"combase.dll");
            if(!combase)combase=LoadLibraryW(L"combase.dll");
            typedef HRESULT (WINAPI *CreateStringFn)(const wchar_t*,UINT,void**);
            CreateStringFn create=(CreateStringFn)GetProcAddress(combase,"WindowsCreateString");
            typedef HRESULT (WINAPI *DeleteStringFn)(void*);
            DeleteStringFn del=(DeleteStringFn)GetProcAddress(combase,"WindowsDeleteString");
            if(!create||!del)hr=ERR_UNSUPPORTED;
            else {
                hr=create(src,(UINT)wcslen(src),&hstr);
                if(SUCCEEDED(hr))hr=((HRESULT (STDMETHODCALLTYPE*)(void*,void*,void*))SLOT(c.mgr,slot))(c.mgr,d,hstr);
                if(hstr)del(hstr);
            }
        }
    }
    RELEASE(d);RELEASE(arr);end(&c);return hr;
}
