#!/bin/bash

IP=$1
PORT=$2
URL=status
SCRIPT_PATH=/home/nagios/monitoring/libexec/CUSTOM/GATEWAY_CHECKS/ONLINE

cd $SCRIPT_PATH
wget http://"$IP":"$PORT"/"$URL"  > OUTPUT 2>&1

count=0
SMSC=ALL

for i in OOREDOO_SMSC_1 OOREDOO_SMSC_2 OOREDOO_SMSC_3
do
SMSC_ONLINE=`grep $i  "$SCRIPT_PATH"/status  | grep online  | awk -F"online"  {'print $2'} | awk -F"rcvd" {'print $1'} | wc -l`
if [ ! $SMSC_ONLINE -eq 1 ]
then
        count=1
        SMSC=$i
        break
fi
done

if [ $count -eq 1 ]
        then
                echo "CRITICAL - SMSC $SMSC is NOT online"
                rm -f  "$SCRIPT_PATH"/status
                exit 2;
        else
                echo "ALL SMSC ARE CONNECTED"
                rm -f  "$SCRIPT_PATH"/status
                exit 0;
fi
