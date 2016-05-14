#! /bin/bash
USER=nagios
PASSWORD=nagios
LIBEXEC=/home/nagios/nagios/libexec
#REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM/GATEWAY_CHECKS/KANNEL_QUEUE
#$LIBEXEC/check_by_ssh -H  $1  -t 90 -q -C "$REMOTE_EXEC/check_kannel_status_queue  $2 $3 "



case $2 in
        gatewayserver)
                REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM/GATEWAY_CHECKS/KANNEL_QUEUE
                $LIBEXEC/check_by_ssh -H  $1  -t 70 -q -C "$REMOTE_EXEC/check_kannel_status_queue  $2 $3 "
        ;;
        kannelserver)
                REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM/KANNEL_CHECKS/KANNEL_QUEUE
                $LIBEXEC/check_by_ssh -H  $1  -t 70 -q -C "$REMOTE_EXEC/check_kannel_status_queue  $2 $3 "
        ;;
        kannel2server)
                REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM/KANNEL2_CHECKS/KANNEL_QUEUE
                $LIBEXEC/check_by_ssh -H  $1  -t 70 -q -C "$REMOTE_EXEC/check_kannel_status_queue  $2 $3 "
        ;;
        smscserver)
                REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM/SMSC_CHECKS/KANNEL_QUEUE
                $LIBEXEC/check_by_ssh -H  $1  -t 70 -q -C "$REMOTE_EXEC/check_kannel_status_queue  $2 $3 "
        ;;


        kannelthreeserver)
                REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM/KANNEL_THREE_CHECKS/KANNEL_QUEUE
                $LIBEXEC/check_by_ssh -H  $1  -t 70 -q -C "$REMOTE_EXEC/check_kannel_status_queue    $2 $3 "
        ;;

        kannelfourserver)
                REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM/KANNEL_FOUR_CHECKS/KANNEL_QUEUE
                $LIBEXEC/check_by_ssh -H  $1  -t 70 -q -C "$REMOTE_EXEC/check_kannel_status_queue    $2 $3 "
        ;;


        kannelfiveserver)
                REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM/KANNEL_FIVE_CHECKS/KANNEL_QUEUE
                $LIBEXEC/check_by_ssh -H  $1  -t 70 -q -C "$REMOTE_EXEC/check_kannel_status_queue    $2 $3 "
        ;;





        *)
                echo "A PROBLEM OCCURED ON LOCAL SCRIPT. PLEASE CHECK"
                exit 2;
        ;;
esac

