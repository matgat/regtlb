#include <iostream>
//#include <fstream>
//#include <sstream>
#include <string>
#include <exception>
//---------------------------------------------------------------------------
#include "unt_OleUts.h" // 'cls_TLBfile'


//---------------------------------------------------------------------------
// Some uts
#include <sys/stat.h>
inline bool file_exists(const std::string& pth) { struct stat buf; return stat(pth.c_str(), &buf)==0; }
std::string extr_file_name(const std::string& pth)
   {
    std::string::size_type i = pth.find_last_of("\\/");
    return (i!=std::string::npos) ? pth.substr(i+1) : pth;
   }


//---------------------------------------------------------------------------
void show_usage()
{
    std::cout << "   Usage:\n";
    std::cout << "   regtlb path-to.tlb     (register tlb file)\n";
    std::cout << "   regtlb -u path-to.tlb  (unregister tlb file)\n";
    std::cout << "   regtlb -i path-to.tlb  (show info about tlb file)\n";
    std::cout << "   regtlb -q {0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}},1,0,0,1  (query tlb path by attributes)\n";
    std::cout << "             <GUID,MajorVer,MinorVer,lcid,syskind>\n";
    std::cout << "       -v: verbose\n";
    std::cout << "       -u: unregister\n";
    std::cout << "       -i: show info of file\n";
    std::cout << "       -q: query by attributes, returns 0 if registered\n";
} // 'show_usage'


//---------------------------------------------------------------------------
enum {RET_OK=0, RET_NOTREG=1, RET_ARGERR=-1, RET_ERR=-2 };
int main( int argc, const char* argv[] )
{
    // (0) Local objects
    std::ios_base::sync_with_stdio(false); // Try to have better performance


    // (1) Command line arguments
    bool verbose = false;
    bool unregister = false;
    bool info = false;
    bool query = false;
    std::string tlb_path, tlb_attr;
    if(verbose) std::cout << "regtlb (" << __DATE__ << ")\n";
    //if(verbose) std::cout << "Running in: " << argv[0] << '\n';

    for( int i=1; i<argc; ++i )
       {
        std::string arg( argv[i] );
        if( arg[0] == '-' )
             {// A command switch
              if( arg.length()!=2 )
                   {
                    std::cerr << "!! Wrong command switch " << arg << '\n';
                    return RET_ARGERR;
                   }
              else if( arg[1]=='v' ) verbose = true;
              else if( arg[1]=='u' ) unregister = true;
              else if( arg[1]=='i' ) info = true;
              else if( arg[1]=='q' ) query = true;
              else {
                    std::cerr << "!! Unknown command switch " << arg << '\n';
                    show_usage();
                    return RET_ARGERR;
                   }
             }
        else {// Not a switch
              if(query)
                   {// Expected an attribute string
                    if( !tlb_attr.empty() )
                       {
                        std::cerr << "!! Too many attributes " << tlb_attr << '\n';
                        show_usage();
                        return RET_ARGERR;
                       }
                    tlb_attr = arg;
                   }
              else {// Expected a path
                    if( !tlb_path.empty() )
                       {
                        std::cerr << "!! Too many paths " << tlb_path << '\n';
                        show_usage();
                        return RET_ARGERR;
                       }
                    tlb_path = arg;
                   }
             }
       } // 'for all cmd args'

    
    // (1') Check
    if( query )
         {
          if( tlb_attr.empty() )
             {
              std::cerr << "!! Specify tlb attributes to query\n";
              show_usage();
              return RET_ARGERR;
             }
         }
    else {// A path is expected
          if( tlb_path.empty() )
             {
              std::cerr << "!! Specify an existing tlb file path\n";
              show_usage();
              return RET_ARGERR;
             }
           else if( !file_exists(tlb_path) )
             {
              std::cerr << "!! File does not exists, check path:\n";
              std::cerr << tlb_path << '\n';
              return RET_ARGERR;
             }
          }


    // (2) Task
    try{
        if( query )
             {
              std::string pth = ole::cls_TLBfile::FindPath(tlb_attr);
              if( pth.empty() )
                   {
                    if(verbose) std::cout << "Not found: " << tlb_attr << '\n';
                    return RET_NOTREG;
                   }
              else {
                    if(verbose) std::cout << "Type library found: " << '\n';
                    std::cout << pth;
                   }
             }
        else {
              std::string tlb_name = extr_file_name(tlb_path);
              ole::cls_TLBfile tlb(tlb_path);
              if( info )
                   {// Just show attributes
                    std::cout << "Attributes of " << tlb_name << ":\n" << tlb.ShowInfo();
                   }
              else if( unregister )
                   {// Want to un-register
                    if( tlb.IsRegistered() )
                       {
                        tlb.Unregister();
                        if(verbose) std::cout << tlb_name << " successfully un-registered\n" << tlb.ShowInfo();
                       }
                    else if(verbose) std::cout << tlb_name << " already un-registered\n";
                   }
              else {// Want to register
                    if( !tlb.IsRegistered() )
                       {
                        tlb.Register();
                        if(verbose) std::cout << tlb_name << " successfully registered\n" << tlb.ShowInfo();
                       }
                    else if(verbose) std::cout << tlb_name << " already registered\n" << tlb.ShowInfo();
                   }
             }
       }
    catch( std::exception& e )
       {
        std::cerr << '\n' << e.what() << '\n';
        return RET_ERR;
       }


    // (3) Finally
    return RET_OK;

} // 'main'
