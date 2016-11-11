#include <cassert> // 'assert'
//#include <exception>
#include <stdexcept> // 'std::runtime_error'
//#include <codecvt> // 'std::wstring_convert'
#include <cctype> // 'std::isspace', 'std::isblank', 'std::isdigit'
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


//---------------------------------------------------------------------------
// Extract hex or decimal integer from string, given current character
int ExtractInt( const std::string& s, int& i )
{
    int val = 0;
    int len = s.length();
    if(i>=len) throw std::runtime_error("No integer, end reached");

    // Detect sign if present
    int sgn = 1; // sign
    if( s[i]=='-' )
       {
        sgn = -1;
        ++i;
        if(i>=len) throw std::runtime_error("No integer, just a minus sign");
       }
    else if( s[i]=='+' )
       {
        ++i;
        if(i>=len) throw std::runtime_error("No integer, just a plus sign");
       }

    // Detect base if present (hexadecimal or octal integer)
    if( s[i]=='0' )
       {
        ++i;
        if(i>=len) return 0;  // Got a zero
        // See char after '0'
        if(s[i]=='x' || s[i]=='X')
           {// Expecting a hexadecimal integer
            val = 0;
            // Get the rest of the hex integer
            while( ++i<len )
               {
                if ( s[i]>='0' && s[i]<='9' ) val = (16*val) + (s[i]-'0');
                else if ( s[i]>='A' && s[i]<='F' ) val = (16*val) + (s[i]-'A'+10);
                else if ( s[i]>='a' && s[i]<='f' ) val = (16*val) + (s[i]-'a'+10);
                else break;
               }
           } // '0x'
        else if( s[i]>='0' && s[i]<='7' )
           {// Expecting an octal integer
            val = s[i] - '0';
            // Get the rest of the octal integer
            while( ++i<len )
               {
                if( s[i]>='0' && s[i]<='7' ) val = (8*val) + (s[i]-'0');
                else if( s[i]=='8' || s[i]=='9' ) throw std::runtime_error("Invalid octal integer at: " + std::to_string(i) + " (" + s[i] + ")");
                else break;
               }
           } // '0'
       }
    else if( s[i]>'0' && s[i]<='9' )
       {// Expecting a decimal integer
        val = s[i] - '0';
        // Get the rest of the decimal integer
        while( ++i<len )
           {
            if( std::isdigit(s[i]) ) val = (10*val) + (s[i]-'0');
            else break;
           }
       }
    else throw std::runtime_error("Invalid integer at: " + std::to_string(i) + " (" + s[i] + ")");
    return sgn * val;
} // 'ExtractInt'


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
// ex. "{0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}}"
std::string GUID2String(const GUID& guid)
{
    std::string s = "{" + mat::ToHex(guid.Data1) + "," +
                          mat::ToHex(guid.Data2) + "," +
                          mat::ToHex(guid.Data3) + ",{";
    for(int k=0; k<8; ++k) s += mat::ToHex(guid.Data4[k]) + ",";
    s[s.length()-1] = '}';
    s += "}";
    return s;
} // 'GUID2String'



//---------------------------------------------------------------------------
// Convert TLIBATTR to string
// ex. "{0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}},1,0,0,1"
std::string Attr2String(const TLIBATTR& a)
{
    std::string s = "{" + mat::ToHex(a.guid.Data1) + "," +
                          mat::ToHex(a.guid.Data2) + "," +
                          mat::ToHex(a.guid.Data3) + ",{";
    for(int k=0; k<8; ++k) s += mat::ToHex(a.guid.Data4[k]) + ",";
    s[s.length()-1] = '}';
    s += "}";
    s += "," + std::to_string(a.wMajorVerNum) + "," + std::to_string(a.wMinorVerNum);
    s += "," + std::to_string(a.lcid);
    s += "," + std::to_string(a.syskind);
    return s;
} // 'Attr2String'



//---------------------------------------------------------------------------
// Convert string to TLIBATTR
// "{0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}},1,0,0,1"
TLIBATTR String2Attr( const std::string& s )
{
    TLIBATTR a;
    int i=0, len=s.length();
    if(i>=len) throw std::runtime_error("Invalid TLIBATTR (empty)");
    if(s[i++]!='{') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    a.guid.Data1 = mat::ExtractInt(s,i);
    if(s[i++]!=',') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    a.guid.Data2 = mat::ExtractInt(s,i);
    if(s[i++]!=',') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    a.guid.Data3 = mat::ExtractInt(s,i);
    if(s[i++]!=',') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    if(i>=len) throw std::runtime_error("Invalid TLIBATTR (truncated Data3)");
    if(s[i++]!='{') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    for(int k=0; k<8; ++k)
       {
        a.guid.Data4[k] = mat::ExtractInt(s,i);
        if(i>=len) throw std::runtime_error("Invalid TLIBATTR (truncated in Data4)");
        else if(k<7 && s[i]!=',') throw std::runtime_error("Invalid TLIBATTR (truncated)");
        else if(k==7 && s[i]!='}') throw std::runtime_error("Invalid TLIBATTR (truncated)");
        ++i;
       }
    if(i>=len) throw std::runtime_error("Invalid TLIBATTR (truncated)");
    if(s[i++]!='}') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    if(i>=len) throw std::runtime_error("Invalid TLIBATTR (truncated after guid)");
    if(s[i++]!=',') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    a.wMajorVerNum = mat::ExtractInt(s,i);
    if(i>=len) throw std::runtime_error("Invalid TLIBATTR (truncated at wMajorVerNum)");
    if(s[i++]!=',') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    a.wMinorVerNum = mat::ExtractInt(s,i);
    if(i>=len) throw std::runtime_error("Invalid TLIBATTR (truncated at wMinorVerNum)");
    if(s[i++]!=',') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    a.lcid = mat::ExtractInt(s,i);
    if(i>=len) throw std::runtime_error("Invalid TLIBATTR (truncated at lcid)");
    if(s[i++]!=',') throw std::runtime_error("Invalid TLIBATTR at: " + std::to_string(i) + " (" + s[i] + ")");
    a.syskind = static_cast<SYSKIND>( mat::ExtractInt(s,i) );
    return a;
} // 'String2Attr'


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
    i_Path = mat::Convert(pth); // Store the path

    HRESULT hr = ::CoInitialize(NULL); // Init COM stuff (single threaded)
    //hr = ::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);  // COINIT_MULTITHREADED
    if( FAILED(hr) ) throw std::runtime_error("Cannot initialize COM: " + HResult2String(hr));

    // Try to load the type library
    try { Load(); }
    catch(std::exception& e) { ::CoUninitialize(); throw; }  // Uninitialize COM stuff and rethrow
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
// std::string pth = cls_TLBfile::FindTLBPath("{0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}},1,0,0,1");
// Query tlb path by attributes <GUID,MajorVer,MinorVer,lcid,syskind>
//const GUID TLB_GUID = {0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}};
//const WORD TLB_MajorVerNum = 1;
//const WORD TLB_MinorVerNum = 0;
//const DWORD TLB_lcid = 0;
//const SYSKIND TLB_syskind = SYS_WIN32; // SYS_WIN16:0,SYS_WIN32:1,SYS_MAC:2,SYS_WIN64:3
//std::string s = ole::cls_TLBfile::FindTLBPath(TLB_GUID,
//                                              TLB_MajorVerNum,
//                                              TLB_MinorVerNum,
//                                              TLB_lcid);
std::string cls_TLBfile::FindPath(const std::string& s)
{
    TLIBATTR a = String2Attr(s);
    BSTR bstr;
    ::QueryPathOfRegTypeLib(a.guid,a.wMajorVerNum,a.wMinorVerNum,a.lcid,&bstr);
    std::wstring wpth = cls_BSTR::Convert(bstr);
    ::SysFreeString(bstr); // Must free the allocated BSTR after use
    return mat::Convert(wpth);
} // 'FindPath'



} // 'nms_Ole' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
