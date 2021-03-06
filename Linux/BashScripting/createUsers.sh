#!/bin/bash

#declares filein as username text file
FILE="Usernames.txt"

# Check for root

if  [ $EUID -ne 0 ]; then
  echo "Must run $0 as root."
  exit 2
fi

# Check for groups and create them if they don't exist

if grep -q staff /etc/group
then
  echo "staff Group exists. Not doing anything."
else
  echo "Creating staff group"
  groupadd staff
  logger "staff group created"
fi

if grep -q visitors /etc/group
then
  echo "visitors Group exists. Not doing anything."
else
  echo "Creating visitors group"
  groupadd visitors
  logger "visitors group created"
fi

#Create an Account for each of the Usernames in the text-file Usernames.txt

if [ ! -f "$FILE" ]then
  echo "Cannot Find the file $filein"
  logger "The file which contains the usernames does not exist as: $filein"
  exit 2
fi

for USER in $FILE
do
  if [ $USER == ,* ]
  then
    useradd -m $USER
    logger "User $USER created"
    echo "User $USER created"
  fi
done

for USER in $FILE
do
    if [[ $USER == ,* ]] && [[ $GROUP == *, ]]
    then
      usermod -a -G $GROUP $USER
      echo "User $USER added to $Group group"
      logger "User $USER added to $Group gorup"
    fi
  done

  echo "Finished"
  logger "Finished"
