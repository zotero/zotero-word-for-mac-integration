#!/bin/bash
cd appscript-*

rm -rf build
export CC=gcc-4.0
export CXX=gcc-4.0
export CFLAGS="-arch ppc -arch i386 -mmacosx-version-min=10.4 -isysroot /Developer/SDKs/MacOSX10.4u.sdk"
export CXXFLAGS="$CFLAGS"
export LDFLAGS="$CFLAGS"
/System/Library/Frameworks/Python.framework/Versions/2.5/bin/python setup.py build
mv build/lib*/aem/ae.so ae-ppc.so

rm -rf build
export CC=gcc-4.2
export CXX=gcc-4.2
export CFLAGS="-arch x86_64 -mmacosx-version-min=10.5 -isysroot /Developer/SDKs/MacOSX10.5.sdk"
export CXXFLAGS="$CFLAGS"
export LDFLAGS="$CFLAGS"
/System/Library/Frameworks/Python.framework/Versions/2.5/bin/python setup.py build
mv build/lib*/aem/ae.so ae-x86.so

AEM=`echo build/lib*/aem`
lipo -create ae-ppc.so ae-x86.so -output $AEM/ae.so
