#!/bin/bash
FILE="/tmp/xx"
if [ ! -e "$FILE" ]
then
  touch "$FILE"
  fi

if [ ! -w "$FILE" ]
then
  echo "I cannot write to the provided location: $FILE"
  logger "Group report cannot write to: $FILE"
  exit 1
fi

grep xx /etc/group 2>&1 > 
$FILElogger "Group report ran and saved to: $FILE"
