#!/bin/bash
BASEDIR=`pwd`

#export CFLAGS='-arch x86_64'
#export CXXFLAGS='-arch x86_64'

export PYTHONHOME="${BASEDIR}/py_install/Python.framework/Versions/2.5"
export PATH=${PYTHONHOME}/bin:$PATH
export PYTHON=${PYTHONHOME}/bin/python
export LD_LIBRARY_PATH=${PYTHONHOME}/lib

# Copy system Python.framework
rm -rf "${BASEDIR}/py_install"
mkdir "${BASEDIR}/py_install"
cp -R /System/Library/Frameworks/Python.framework "${BASEDIR}/py_install/Python.framework"
rm -rf "${BASEDIR}/py_install/Python.framework/Versions/2."[^5]
rm "${BASEDIR}/py_install/Python.framework/Versions/Current"
ln -s 2.5 "${BASEDIR}/py_install/Python.framework/Versions/Current"

# Build
rm -rf "${BASEDIR}/pythonext-"*.xpi
for arch in "i386" "x86_64"; do
	rm -rf "${BASEDIR}/xulrunner-sdk"
	rm -rf "${BASEDIR}/pythonext@mozdev.org-$arch"
	cp -R "${BASEDIR}/xulrunner-sdk-$arch" "${BASEDIR}/xulrunner-sdk"
	
	if [ $arch == "i386" ]; then
		export CFLAGS="-arch i386 -mmacosx-version-min=10.5 -isysroot /Developer/SDKs/MacOSX10.5.sdk"
	else
		export CFLAGS="-arch x86_64"
	fi
	
	export CXXFLAGS="$CFLAGS"
	export LDFLAGS="-F${BASEDIR}/py_install $CFLAGS"
	${PYTHON} build_pythonext.py
	
	mv "${BASEDIR}/pythonext@mozdev.org" "${BASEDIR}/pythonext@mozdev.org-$arch"
	mv "${BASEDIR}/pythonext-"*Darwin_x86-*.xpi "${BASEDIR}/pythonext-Darwin_"$arch.xpi
	rm -rf "${BASEDIR}/xulrunner-sdk"
done

# Unify
cp -r "${BASEDIR}/pythonext@mozdev.org-i386" "${BASEDIR}/pythonext@mozdev.org"
for lib in "components/libpybootstrap.dylib" \
           "pylib/libpyloader.dylib" \
           "pylib/libpyxpcom.dylib" \
           "pylib/xpcom/_xpcom.so"
do
	lipo -create "${BASEDIR}/pythonext@mozdev.org-i386/$lib" \
		"${BASEDIR}/pythonext@mozdev.org-x86_64/$lib" \
		-output "${BASEDIR}/pythonext@mozdev.org/$lib"
done

cd "${BASEDIR}/pythonext@mozdev.org"


# Fix ABI
perl -pi -e 's|<em:targetPlatform>Darwin_x86-gcc3</em:targetPlatform>|<em:targetPlatform>Darwin_x86-gcc3</em:targetPlatform>\n        <em:targetPlatform>Darwin_x86_64-gcc3</em:targetPlatform>|' install.rdf

zip -r ../pythonext-Darwin_universal.xpi * -x '*CVS/*' '*.svn/*' '*.hg/*'