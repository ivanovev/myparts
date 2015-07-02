#!/bin/bash

if [ $# == 0 ]; then
    echo "wrong # of args"
    exit 1
fi

#path=`pwd`
#name=`basename $path`
MYFILES=myparts.py \
        META-INF/manifest.xml
OXTFILE=/tmp/myparts.oxt

make_oxt()
{
    rm -f $OXTFILE
    zip -r $OXTFILE $(MYFILES)
}

commit()
{
    rm -rf $(mydir)
    n=`find . -name *swp | wc -l`
    if [ $n -ne 0 ]; then
        echo 'close all vim instances, plz...'
        exit 1
    fi
    read -p "Commit message: " -e msg
    git add ./* && git commit -m "$msg" && git push
}

case $1 in
'oxt') make_oxt
        ;;
'commit') commit
        ;;
'stats') find . -name '*.c' -print0 | xargs -0 cat | egrep -v '^[ \t]*$' | wc
        ;;
*)      echo 'invalid option'
        ;;
esac

