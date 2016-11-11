  /---------------------------------------------------\
        Mat’s Type Library registration utility
          ©2016 Matteo Gattanini
  \---------------------------------------------------/
   Created specifically for api8070.tlb


  /---------------------------------------------------\
    Versioning
  \---------------------------------------------------/
   Hg repository is in:
   https://matgat@bitbucket.org/matgat/regtlb


  /---------------------------------------------------\
    Usage
  \---------------------------------------------------/
    mpp -v -m .ncs=.fst -i defvar.def -f *.ncs
    mpp -x -i defvar.h "in.nc" - "out.ncs"
    mpp -f -i defvar.h -i iomap.h -i messages.h Interface.xml

        -v: verbose
        -f: force overwrite
        -x: invert the dictionary
        -i: include definitions file
        -c: declare line comment char
        -m: map automatic output extension

