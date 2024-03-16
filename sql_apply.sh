#!/bin/bash


# Docker command: docker run -v {local/path/to/files}:/wo --network {NETWORK NAME} -it mcr.microsoft.com/mssql-tools

GREEN='\033[0;32m'
RED='\033[0;31m'
NC='\033[0m' # No Color

read -p 'Enter Server: ' servername #db,1433 or '{container name},{port}'
read -p 'Enter User: ' username # sa
read -sp 'Enter Password: ' password

runSqlFile() {
  if [ "$1" != "" ]; then
    echo -en "${1}"
    sqlcmd -m 1 -b -h -1 -r 0 -S ${servername} -U ${username} -P ${password} -i ${1} -o ${1}.out < /dev/null > /dev/null
    if [ -s ${1}.out ] && grep -q -E "Msg|Level|State|Error:|Warning:" "$1.out"; then
        echo -e "$\r${RED}FAILED: ${1}${NC}"
        mv "${1}.out" "${1}.err"
    else
        echo -e "$\r${GREEN}Success: ${1}${NC}"
        rm "${1}.out"
    fi
  else
    echo -e "${RED}ERROR: Empty Filename${NC}"
  fi

}
echo
echo ..................................
echo Start
echo ..................................
echo "Note: Be Patient, this Patch will take a considerable amount of time to run (wait) "
echo ..................................

echo
echo ..................................
echo Loading Data from lst file...
echo ..................................
echo

while IFS= read -r line
do
  runSqlFile $line
done < data_files.lst

echo
echo ..................................
echo Verifying the Changes...
echo ..................................
echo
runSqlFile verify.sql

echo
echo ..................................
echo Checking output files for Errors
echo ..................................
echo
if [ -f *.sql.err ]; then
    echo -e "$\r${RED}Error Files exist!";
    ls *.sql.err
    echo -e "\r${NC}"
fi

echo ..................................
echo End
echo ..................................
