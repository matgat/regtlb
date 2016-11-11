#ifndef UNT_OLEUTS_H
#define UNT_OLEUTS_H
/*  ---------------------------------------------
    unt_OleUts.h
    ©2016 matteo.gattanini@gmail.com

    OVERVIEW
    ---------------------------------------------
    A unit that collects some utilities

    LICENSES
    ---------------------------------------------
    Use and modify freely

    EXAMPLE OF USAGE:
    ---------------------------------------------
    #include "unt_OleUts.h" // 'cls_TLBfile'

    DEPENDENCIES:
    ---------------------------------------------     */
    #include <string>




namespace nms_Ole //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::
{
    std::string HResult2String( const long /* HRESULT */ );
    


/////////////////////////////////////////////////////////////////////////////
// Wrapper to Type Library file for registration facilities
class cls_TLBfile ///////////////////////////////////////////////////////////
{
 public:
    // . Public methods
    cls_TLBfile(const std::string& pth);
    ~cls_TLBfile();

    void Register(); // Register Type Library
    void Unregister(); // Unregister Type Library

    const std::wstring& Path() const { return i_Path; }; // Stored path
    bool IsRegistered() { try{ return QueryPath().length()>1; } catch(...){ return false; } }; // Tell if already registered
    std::string ShowInfo(); // Show Library attributes
    static std::string FindPath(const std::string& s); // Retrieves the path of a registered tlb from attributes

 private:
    class ITypeLib* pTlb;
    class tagTLIBATTR* pTlbAttr;
    std::wstring i_Path;

    void Load(); // Load the DLL library
    void Query(); // Retrieve library info
    std::wstring QueryPath(); // Retrieves the path of this type library

}; // 'cls_TLBfile'




} // 'nms_Ole' ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
namespace ole = nms_Ole; // a short alias

#endif // 'UNT_OLEUTS_H'
