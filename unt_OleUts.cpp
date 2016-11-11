#include <cassert> // 'assert'
//#include <exception>
#include <stdexcept> // 'std::runtime_error'
//#include <codecvt> // 'std::wstring_convert'
#include <windows.h> // Windows stuff
#include <objbase.h> // COM stuff,	oleauto.h
//---------------------------------------------------------------------------
#include "unt_OleUts.h"




//:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
namespace mat //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
{
  //inline T_BYTE HiByte(const T_WORD w){return T_BYTE((w>>8) & 0x00FF);}
  //inline T_BYTE LoByte(const T_WORD w){return T_BYTE(w & 0x00FF);}
  //inline T_WORD HiWord(const T_DWORD dw){return T_WORD((dw>>16) & 0x0000FFFF);}
  //inline T_WORD LoWord(const T_DWORD dw){return T_WORD(dw & 0x0000FFFF);}

  //template<T_BYTE> std::string ToHex(const T_BYTE b) {char s[2]={hexdig[(b>>4)&0x0F],hexdig[b&0x0F]}; return std::string(s,2);}
  //template<T_WORD> std::string ToHex(const T_WORD w) {char s[4]={hexdig[(w>>12)&0x0F],hexdig[(w>>8)&0x0F],hexdig[(w>>4)&0x0F],hexdig[w&0x0F]}; return std::string(s,4);}

  //-------------------------------------------------------------------------
  // Convert to hexadecimal string
  template<typename T> std::string ToHex(const T v)
     {
      const char hexdig[] = "0123456789ABCDEF";
      const int n = 2*sizeof(T);
      char s[n];
      for(int i=0; i<n; ++i) s[i] = hexdig[(v>>(4*(n-1-i))) & 0xF];
      return std::string(s,n);
     }

  //-------------------------------------------------------------------------
  // Convert UTF-8 string to UTF-16
  std::wstring Convert(const std::string& s)
     {
      // C++14 style
      //return std::wstring_convert<std::codecvt_utf8<char>>().to_bytes(s);
      //return std::wstring_convert<std::codecvt_utf16<wchar_t>>().from_bytes(s);

      // Windows style, to UCS16
      //int ws_length = ::MultiByteToWideChar(CP_UTF8, 0, s.data(), s.length(), ws.data(), 0);
      //std::wstring s(ws_length, '\0');
      //::MultiByteToWideChar(CP_UTF8, 0, s.data(), s.length(), ws.data(), ws.length());

      // Poor man's conversion
      return std::wstring(s.begin(), s.end());
     }

  //-------------------------------------------------------------------------
  // Convert UTF-16 string to UTF-8
  std::string Convert(const std::wstring& ws)
     {
      // C++14 style
      //return std::wstring_convert<std::codecvt_utf8<char>>().to_bytes(s);
      //return std::wstring_convert<std::codecvt_utf16<wchar_t>>().from_bytes(s);

      // Windows style, to UCS16
      //int s_length = ::WideCharToMultiByte(CP_UTF16, 0, ws.data(), ws.length(), s.data(), 0);
      //std::string s(s_length, '\0');
      //::WideCharToMultiByte(CP_UTF16, 0, ws.data(), ws.length(), s.data(), s.length());

      // Poor man's conversion
      return std::string(ws.begin(), ws.end());
     }

} // 'mat' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::



//:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
namespace nms_Ole //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::
{



  ///////////////////////////////////////////////////////////////////////////
  // Wrapper to the fucking BSTR
  class cls_BSTR ////////////////////////////////////////////////////////////
  {
   public:
      // . Public methods
      cls_BSTR(const std::string& s) { i_bstr = Convert(s); }
      cls_BSTR(const std::wstring& s) { i_bstr = Convert(s); }
      ~cls_BSTR() { ::SysFreeString(i_bstr); }

      operator BSTR() const { return i_bstr; }
      operator std::wstring() const { return Convert(i_bstr); }


      //---------------------------------------------------------------------
      static std::wstring Convert(BSTR bstr)
         {
          return (bstr != nullptr) ? std::wstring(bstr, ::SysStringLen(bstr)) : L"";
          //long wslen = ::SysStringLen(bstr);
          //int len = ::WideCharToMultiByte(CP_ACP, 0, (wchar_t*)bstr, wslen, NULL, 0, NULL, NULL);
          //std::wstring s(len, '\0');
          //len = ::WideCharToMultiByte(CP_ACP, 0, (wchar_t*)bstr, wslen, &s[0], len, NULL, NULL );
          //return s;
         }

      //---------------------------------------------------------------------
      static BSTR Convert(const std::wstring& ws)
         {
          return ws.empty() ? nullptr : ::SysAllocStringLen(ws.data(), ws.size());
          //int wslen = ::MultiByteToWideChar(CP_ACP, 0, ws.data(), ws.length(), NULL, 0);
          //BSTR bs = ::SysAllocStringLen(NULL, wslen);
          //::MultiByteToWideChar(CP_ACP, 0, ws.data(), ws.length(), bs, wslen);
          //return bs; // remember to 'SysFreeString' on this!
         }

      //---------------------------------------------------------------------
      static BSTR Convert(const std::string& s)
         {
          std::wstring ws = mat::Convert(s);
          return ws.empty() ? nullptr : ::SysAllocStringLen(ws.data(), ws.size());
         }

   private:
      BSTR i_bstr;

  }; // 'cls_BSTR'


//---------------------------------------------------------------------------
// Convert GUID to string
// ex. {0x01751AD4,0x743E,0x4578,{0x9B,0x56,0x96,0x35,0x1D,0x18,0x6D,0x01}}
std::string GUID2String(const GUID& guid)
{
    std::string s = "{" + mat::ToHex(guid.Data1) + "," +
                          mat::ToHex(guid.Data2) + "," +
                          mat::ToHex(guid.Data3) + ",{";
    for(int i=0; i<8; ++i) s += mat::ToHex(guid.Data4[i]) + ",";
    s[s.length()-1] = '}';
    s += "}";
    return s;
} // 'GUID2String'



//---------------------------------------------------------------------------
// Get a string from an HRESULT (these macros are defines in 'winerror.h')
std::string HResult2String(const HRESULT hr)
{
#ifndef E_RPCSVR_UNAVAILABLE
  #define  E_RPCSVR_UNAVAILABLE  0x800706BAL
#endif

    switch (hr)
       {
        case E_OUTOFMEMORY          : return "out of memory!!";
        case E_INVALIDARG           : return "one or more arguments are invalid!!";
        case REGDB_E_CLASSNOTREG    : return "class not registered!!";
        case REGDB_E_WRITEREGDB     : return "registry write error!!";
        case E_UNEXPECTED           : return "unexpected error occurred!!";
        case CLASS_E_NOAGGREGATION  : return "class cannot be part of an aggregate!!";
        case CO_S_NOTALLINTERFACES  : return "failed to retrieve some of the interfaces!!";
        case E_NOINTERFACE          : return "failed to retrieve all the interfaces!!";
        case TYPE_E_IOERROR         : return "cannot write the file!!";
        case TYPE_E_REGISTRYACCESS  : return "cannot open the register!!";
        case TYPE_E_INVALIDSTATE    : return "cannot open the type library!!";
        case TYPE_E_CANTLOADLIBRARY : return "cannot load the dll/type library!!";
        case TYPE_E_UNKNOWNLCID     : return "cannot find the passed LCID in the OLE-supported DLLs!!";
        case TYPE_E_INVDATAREAD     : return "cannot read the file!!";
        case TYPE_E_UNSUPFORMAT     : return "the type library has an older format!!";
        //case E_OUT_OF_MEMORY        : return "out of memory";
        case RPC_E_TOO_LATE         : return "CoInitializeSecurity has already been called";
        case RPC_E_NO_GOOD_SECURITY_PACKAGES : return "asAuthSvc par was not NULL, and none of the auth services in the list could be registered. Check the results in asAuthSvc";
        case S_FALSE                : return "nothing to do";
        case E_ACCESSDENIED         : return "access denied";
        case CO_E_CLASSSTRING       : return "the registered CLSID for the ProgID is invalid";
        case E_RPCSVR_UNAVAILABLE   : return "rpc server unavailable";
        case CO_E_NOTINITIALIZED    : return "CoInitialize has not been called";
        case S_OK                   : return "success";
        default                     : return "unknown retcode "+std::to_string(hr)+"!";
       }
} // 'HResult2String'







/////////////////////////////////////////////////////////////////////////////
///////////////////////////////// cls_TLBfile ///////////////////////////////


//---------------------------------------------------------------------------
// Constructor
cls_TLBfile::cls_TLBfile(const std::string& pth) : pTlb(nullptr), pTlbAttr(nullptr)
{
    HRESULT hr = ::CoInitialize(NULL); // Init COM stuff (single threaded)
    //hr = ::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);  // COINIT_MULTITHREADED
    if( FAILED(hr) ) throw std::runtime_error("Cannot initialize COM: " + HResult2String(hr));

    // Try to load the type library
    try { Load(); }
    catch(std::exception& e)
       {
        ::CoUninitialize(); // Uninitialize COM stuff
        throw;
       }

    i_Path = mat::Convert(pth); // Store the path
}

//---------------------------------------------------------------------------
// Destructor
cls_TLBfile::~cls_TLBfile()
{
    if(pTlbAttr) pTlb->ReleaseTLibAttr(pTlbAttr);
    pTlb->Release();
    ::CoUninitialize(); // Uninitialize COM stuff
}

//---------------------------------------------------------------------------
// Register a type library
void cls_TLBfile::Register()
{
    cls_BSTR pth(i_Path);
    HRESULT hr = ::RegisterTypeLib(pTlb, pth, NULL); // Register the library
    // Check if registration was successful
    if(FAILED(hr)) throw std::runtime_error("Cannot register the type library: " + HResult2String(hr));
} // 'Register'

//---------------------------------------------------------------------------
// Unregister a type library
void cls_TLBfile::Unregister()
{
    if(!pTlbAttr) Query();
    // Unregister the library
    HRESULT hr = ::UnRegisterTypeLib(pTlbAttr->guid, pTlbAttr->wMajorVerNum, pTlbAttr->wMinorVerNum, pTlbAttr->lcid, pTlbAttr->syskind);
    // Check if unregistration was successful
    if( FAILED(hr) )  // && hr!=TYPE_E_REGISTRYACCESS
       {
        throw std::runtime_error("Cannot unregister the type library: " + HResult2String(hr));
       }
} // 'Unregister'

//---------------------------------------------------------------------------
// Convert Library attributes to string
std::string cls_TLBfile::ShowInfo()
{
    if(!pTlbAttr) Query();
    std::string s;
    if( pTlbAttr )
       {
        s = "Path: " + mat::Convert(i_Path) + "\n" "GUID: " + GUID2String(pTlbAttr->guid) + "\n";
        s += "Version: " + std::to_string(pTlbAttr->wMajorVerNum) + " . " + std::to_string(pTlbAttr->wMinorVerNum) + "\n"; // (WORD)
        s += "lcid: " + std::to_string(pTlbAttr->lcid) + "\n"; // (DWORD)
        s += "sys: ";
        switch(pTlbAttr->syskind)
           {
          #if SYS_WIN32 == SYS_WIN64
            case SYS_WIN32 : s += "win32/64"; break;
          #else
            case SYS_WIN32 : s += "win32"; break;
            case SYS_WIN64 : s += "win64"; break;
          #endif
            case SYS_WIN16 : s += "win16"; break;
            case SYS_MAC   : s += "mac"; break;
            default        : s += "unknown";
           }
       }
    return s;
} // 'ShowInfo'





//---------------------------------------------------------------------------
// Retrieve library info (if already registered)
void cls_TLBfile::Query()
{
    if(pTlbAttr) pTlb->ReleaseTLibAttr(pTlbAttr);
    HRESULT hr = pTlb->GetLibAttr(&pTlbAttr);
    if( FAILED(hr) )
       {
        pTlbAttr = 0;
        throw std::runtime_error("Cannot get attributes of type library: " + HResult2String(hr));
       }
} // 'Query'


//---------------------------------------------------------------------------
// Load the library
void cls_TLBfile::Load()
{
    if(pTlb) pTlb->Release();
    cls_BSTR pth(i_Path);
    HRESULT hr = ::LoadTypeLibEx(pth, REGKIND_NONE, &pTlb);  // REGKIND_REGISTER
    if(FAILED(hr)) throw std::runtime_error("Cannot load the type library: " + HResult2String(hr));
}


//---------------------------------------------------------------------------
// Retrieves the path of this type library
std::wstring cls_TLBfile::QueryPath()
{
    if(!pTlbAttr) Query();
    BSTR bstr;
    ::QueryPathOfRegTypeLib(pTlbAttr->guid,pTlbAttr->wMajorVerNum,pTlbAttr->wMinorVerNum,pTlbAttr->lcid,&bstr);
    std::wstring wpth = cls_BSTR::Convert(bstr);
    ::SysFreeString(bstr); // Must free the allocated BSTR after use
    return wpth;
} // 'QueryPath'

//---------------------------------------------------------------------------
// Retrieves the path of an already registered type library
// std::string pth = cls_TLBfile::FindTLBPath("{0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}},1,0,0");
std::string cls_TLBfile::FindPath(const std::string& s)
{
    TLIBATTR a;
    // TODO: parse string
    BSTR bstr;
    ::QueryPathOfRegTypeLib(a.guid,a.wMajorVerNum,a.wMinorVerNum,a.lcid,&bstr);
    std::wstring wpth = cls_BSTR::Convert(bstr);
    ::SysFreeString(bstr); // Must free the allocated BSTR after use
    return mat::Convert(wpth);
} // 'FindPath'



    //
    //(query tlb path by attributes <GUID,MajorVer,MinorVer,lcid,syskind>


    //const GUID TLB_GUID = {0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}};
    //const WORD TLB_MajorVerNum = 1;
    //const WORD TLB_MinorVerNum = 0;
    //const DWORD TLB_lcid = 0;
    //const SYSKIND TLB_syskind = SYS_WIN32; // SYS_WIN16:0,SYS_WIN32:1,SYS_MAC:2,SYS_WIN64:3
    //api8070RegPth = mat::cls_TLBfile::FindTLBPath(cnc8070::TLB_GUID,
    //                                              cnc8070::TLB_MajorVerNum,
    //                                              cnc8070::TLB_MinorVerNum,
    //                                              cnc8070::TLB_lcid);


} // 'nms_Ole' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
