#!/bin/bash

IP=$1
PORT=$2
URL=status
SCRIPT_PATH=/home/nagios/monitoring/libexec/CUSTOM/GATEWAY_CHECKS/KANNEL_QUEUE


cd $SCRIPT_PATH
/usr/bin/wget   http://"$IP":"$PORT"/"$URL" -O $URL  > $SCRIPT_PATH/OUTPUT 2>&1

QUEUE_SIZE=`grep SMS $SCRIPT_PATH/status  | grep received | grep queued | awk -F"queued" {'print $2'} | awk -F"(" {'print $2'}`
QUEUE_SIZE_PREVIOUS=`cat $SCRIPT_PATH/PREVIOUS_VALUE  | awk {'print $1'}`
##QUEUE_SIZE=100

if [ $QUEUE_SIZE -lt 70 ]
then
        echo "KANNEL QUEUE IS LESS THAN 70.CURRENT VALUE: $QUEUE_SIZE"
        echo $QUEUE_SIZE > $SCRIPT_PATH/PREVIOUS_VALUE
        rm -f $SCRIPT_PATH/status
        exit 0;
fi

if [ $QUEUE_SIZE -gt 70 -a $QUEUE_SIZE_PREVIOUS -gt 70 ]
then
        echo "WARNING - KANNEL QUEUE SIZE WAS FOUND GREATER THAN 70 ON 2 SEQUENTAL CHECKS. CURRENT VALUE: $QUEUE_SIZE"
        rm -f $SCRIPT_PATH/status
        echo $QUEUE_SIZE > $SCRIPT_PATH/PREVIOUS_VALUE
        exit 1;

fi


if [ $QUEUE_SIZE -gt 90 -a $QUEUE_SIZE_PREVIOUS -gt 90 ]
then
        echo "CRITICAL - KANNEL QUEUE SIZE WAS FOUND GREATER THAN 90 ON 2 SEQUENTAL CHECKS.CURRENT VALUE: $QUEUE_SIZE"
        echo $QUEUE_SIZE > $SCRIPT_PATH/PREVIOUS_VALUE
        rm -f $SCRIPT_PATH/status
        exit 1;

fi




if [ $QUEUE_SIZE -gt 98  ]
then
        echo "CRITICAL - KANNEL QUEUE SIZE IS FULL.CURRENT VALUE: $QUEUE_SIZE"
        echo $QUEUE_SIZE > $SCRIPT_PATH/PREVIOUS_VALUE
        rm -f $SCRIPT_PATH/status
        exit 2
fi


echo "KANNEL QUEUE IS OK. CURRENT VALUE: $QUEUE_SIZE"
echo $QUEUE_SIZE > $SCRIPT_PATH/PREVIOUS_VALUE
exit 0

rm -f $SCRIPT_PATH/status
