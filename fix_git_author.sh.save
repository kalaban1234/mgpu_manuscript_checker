#!/bin/bash

git filter-branch --env-filter '

OLD_EMAIL="kalaban1234@example.com"
CORRECT_NAME="kalaban1234"
CORRECT_EMAIL="kalaban1234@gmail.com"

if [ "$GIT_COMMITTER_EMAIL" = "$OLD_EMAIL" ]
then
    export GIT_COMMITTER_NAME="$CORRECT_NAME"
    export GIT_COMMITTER_EMAIL="$CORRECT_EMAIL"
fi
if [ "$GIT_AUTHOR_EMAIL" = "$OLD_EMAIL" ]
then
    export GIT_AUTHOR_NAME="$CORRECT_NAME"
    export GIT_AUTHOR_EMAIL="$CORRECT_EMAIL"
fi
' --tag-name-filter cat -- --branches --tags

echo "Author info rewritten to $CORRECT_NAME <$CORRECT_EMAIL> in all commits."
