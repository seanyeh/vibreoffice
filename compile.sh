#!/bin/sh
#
# "Compile" BASIC macro file to LibreOffice-compatible xba file

# compile SRC DESTINATION
compile() {
	# Escape XML &<>'"
	src=`sed "s/\&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g; s/'/\&apos;/g; s/\"/\&quot;/g" "$1"`

	xbafile="$2"
    name="`basename -s \".xba\" "$xbafile"`"

	XBA_TEMPLATE='<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n<script:module xmlns:script="http://openoffice.org/2000/script" script:name="%s" script:language="StarBasic">\n%s\n</script:module>'

	printf "$XBA_TEMPLATE" "$name" "$src" > "$xbafile"
}

if [ "$#" -ne 2 ]; then
    echo "Usage: $0 SOURCE DESTINATION"
else
    compile "$1" "$2"
fi
