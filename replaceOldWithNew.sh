#!/bin/sh
for file in $(find |grep .sql); do
	echo "Processing $file"
	sed -i -e 's/\[old\]/\[new\]/g' $file
	sed -i -e 's/\[OLD\]/\[new\]/g' $file
	sed -i -e 's/old\./new\./g' $file
	sed -i -e "s:N'old'.:N'new'.:" $file
done
