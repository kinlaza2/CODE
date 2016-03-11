#!/bin/bash

IP=$1
PORT=$2
URL=status
FLAG=$3
SCRIPT_PATH=/home/nagios/monitoring/libexec/CUSTOM/KANNEL_CHECKS/DLR_CHECK
SMSCS="SMSCS: "
LIST_OF_SMSC=KUWAIT_MO_MT
status=0
EXIT_PAGE=status
QUEUE_LIMIT_WARNING=30000
QUEUE_LIMIT_CRITICAL=90000

###################################### CHECK NUMBER OF PARAMETERS ##############################################################################

############### TBA  ----            TO BE ADDED


#re='^[0-9]+$'
#if ! [[ $1 =~ $re ]] ; then
#   echo "error: Not a number" >&2; exit 1
#fi

#exit 0;






################################################################################################################################################
###################################### PROGRAM START ###########################################################################################


cd $SCRIPT_PATH
wget http://"$IP":"$PORT"/"$URL" > OUTPUT 2>&1


case $FLAG in
DLR_QUEUE)
    #################################### CHECK THAT DLR QUEUE IS NOT BIG #########################################################################
    PENDING_MESSAGES_ONQUEUE=`grep queued $SCRIPT_PATH/$EXIT_PAGE  | grep "using mysql storage"  | awk -F "DLR:" {'print $2'} | awk {'print $1'}`
    re='^[0-9]+$'
    if ! [[ $PENDING_MESSAGES_ONQUEUE =~ $re ]] ; then
        rm -f  "$SCRIPT_PATH"/$EXIT_PAGE
        echo "ERROR OCCUREAD. DLR QUEUE NUMBER IS NOT CORRECT" >&2; exit 2
    fi

    if [ $PENDING_MESSAGES_ONQUEUE -gt $QUEUE_LIMIT_CRITICAL ]
    then
        rm -f  "$SCRIPT_PATH"/$EXIT_PAGE
        echo -e "CRITICAL - PROBLEM ON QUEUE - MESSAGES ON QUEUE ARE: $PENDING_MESSAGES_ONQUEUE WHEN LIMIT IS $QUEUE_LIMIT_WARNING"
        exit 2;
    else
        if    [ $PENDING_MESSAGES_ONQUEUE -gt $QUEUE_LIMIT_WARNING ]
        then
            rm -f  "$SCRIPT_PATH"/$EXIT_PAGE
            echo -e "WARNING - PROBLEM ON QUEUE - MESSAGES ON QUEUE ARE: $PENDING_MESSAGES_ONQUEUE WHEN LIMIT IS $QUEUE_LIMIT_WARNING"
            exit 1;
        else
            rm -f  "$SCRIPT_PATH"/$EXIT_PAGE
            echo -e "OK - DLR QUEUE CONTAINS $PENDING_MESSAGES_ONQUEUE MESSAGES ON QUEUE WHEN LIMIT IS $QUEUE_LIMIT_WARNING "
            exit 0;
        fi
    fi
;;

    #################################### CHECK THAT WE HAVE RECEIVED AT LEAST ONE DLR FOR SOME TIME ##############################################
DLR_RECEIVE)
################################# TBA  ----           TO BE ADDED WARNING LEVEL

    for i in $LIST_OF_SMSC
    do
        if [ ! -f "$SCRIPT_PATH"/PREVIOUS_VALUE_"$i" ]
        then
                touch "$SCRIPT_PATH"/PREVIOUS_VALUE_"$i"
                echo 0 > "$SCRIPT_PATH"/PREVIOUS_VALUE_"$i"
        fi
    done


    for i in $LIST_OF_SMSC
    do
    SMSC_CURRENT_VALUE=`cat  "$SCRIPT_PATH"/$EXIT_PAGE | grep DLR | grep received | awk {'print $3'} | replace "," ""`
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
        echo "CRITICAL - FOR $SMSCS THE NUMBER OF RECEIVED DLRs HAS NOT BEEN CHANGED FROM PREVIOUS CHECK ($SMSC_CURRENT_VALUE)."
        rm -f  "$SCRIPT_PATH"/$EXIT_PAGE
        exit 2;
    else
        echo "ALL SMSC HAVE RECEIVED AT LEAST ONE DLR  FROM PREVIOUS CHECK."
        rm -f  "$SCRIPT_PATH"/$EXIT_PAGE
        exit 0;
    fi
  ;;

  *)
    echo "UNEXPECTED ERROR"
    exit 2;
esac

