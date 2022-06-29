#!/bin/bash

function usage {
    echo "Usage: $(basename $0) <filename> <commit_message>"    
}

function die {
    declare MSG="$@"
    echo -e "$0: Error: $MSG">&2
    exit 1
}

(( "$#" == 1 )) || die "Wrong arguments.\n\n$(usage)"

FILE="${1}"
COMMIT_MESSAGE="autocommit"

[ -f "${FILE}" ] || die "File $FILE does not exist"

echo -n "adding $FILE to git..."
git add "${FILE}" || die "git add $FILE has failed."
echo done

echo "commiting $file to git..."
git commit -m "$COMMIT_MESSAGE" || die "git commit has failed."

echo "pushing to origin..."
git push origin || die "git push origin has failed"

exit 0