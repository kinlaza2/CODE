#!/bin/bash

IP=$1
PORT=$2
URL=status
SCRIPT_PATH=/home/nagios/monitoring/libexec/CUSTOM/KANNEL_CHECKS/DLR_CHECK
SMSCS="SMSCS: "
status=0

cd $SCRIPT_PATH
wget http://"$IP":"$PORT"/"$URL" > OUTPUT 2>&1

#################################### CHECK THAT DLR QUEUE IS NOT BIG #########################################################################




#################################### CHECK THAT WE HAVE RECEIVED AT LEAST ONE DLR FOR SOME TIME ##############################################
for i in KUWAIT_MO_MT
do
        if [ ! -f "$SCRIPT_PATH"/PREVIOUS_VALUE_"$i" ]
        then
                touch "$SCRIPT_PATH"/PREVIOUS_VALUE_"$i"
                echo 0 > "$SCRIPT_PATH"/PREVIOUS_VALUE_"$i"
        fi
done


for i in KUWAIT_MO_MT
do
SMSC_CURRENT_VALUE=`cat  "$SCRIPT_PATH"/status | grep DLR | grep received | awk {'print $3'} | replace "," ""`
SMSC_PREVIOUS_VALUE=`cat "$SCRIPT_PATH"/PREVIOUS_VALUE_"$i"`



if [ ! $SMSC_CURRENT_VALUE -gt $SMSC_PREVIOUS_VALUE ]
then
                #echo "CRITICAL - SMSC $i : $THE NUMBER OF RECEIVED DLRs HAS NOT BEEN CHANGED FROM PREVIOUS CHECK."
                status=1
                SMSCS=`echo $SMSCS " " $i`
fi


echo $SMSC_CURRENT_VALUE > "$SCRIPT_PATH"/PREVIOUS_VALUE_"$i"

done

if [ $status -eq 1 ]
then
        echo "CRITICAL - FOR $SMSCS THE NUMBER OF RECEIVED DLRs HAS NOT BEEN CHANGED FROM PREVIOUS CHECK."
        rm -f  "$SCRIPT_PATH"/status
        exit 2;
else
        echo "ALL SMSC HAVE RECEIVED AT LEAST ONE DLR  FROM PREVIOUS CHECK."
        rm -f  "$SCRIPT_PATH"/status
        exit 0;
fi

