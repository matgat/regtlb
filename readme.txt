  /---------------------------------------------------\
        Mat’s Type Library registration utility
          ©2016 Matteo Gattanini
  \---------------------------------------------------/
   win32 console application created specifically to
   manage api8070.tlb registration


  /---------------------------------------------------\
    Versioning
  \---------------------------------------------------/
   Hg repository is in:
   https://matgat@bitbucket.org/matgat/regtlb


  /---------------------------------------------------\
    Usage
  \---------------------------------------------------/
    Usage:
    regtlb path-to.tlb     (register tlb file)
    regtlb -u path-to.tlb  (unregister tlb file)
    regtlb -i path-to.tlb  (show info about tlb file)
    regtlb -q {0x714D710F,0x27B7,0x42BA,{0xA5,0x88,0x91,0xAD,0xC7,0xFF,0x4B,0x47}},1,0,0  (query tlb path by attributes)
              <GUID,MajorVer,MinorVer,lcid>
        -v: verbose
        -u: unregister
        -i: show info of file
        -q: query by attributes, returns 0 if registered

