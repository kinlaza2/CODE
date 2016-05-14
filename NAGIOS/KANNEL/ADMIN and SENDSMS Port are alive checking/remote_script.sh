CHECK_TCP_PATH=/home/nagios/monitoring/libexec
ADMIN_PORT=$2
SENDSMS_PORT=$3
HOST=$1

###### PROGRAM ####################
ADMIN_PORT_RESPONSE=` $CHECK_TCP_PATH/check_tcp -H $HOST -p $ADMIN_PORT`
if [ ` echo $ADMIN_PORT_RESPONSE | grep 'TCP OK'  | wc -l` -eq 1  ]
then
        ISADMIN_PORT_RESPOND=1  #we are ok
else
        ISADMIN_PORT_RESPOND=0
fi

SENDSMS_PORT_RESPONSE=` $CHECK_TCP_PATH/check_tcp -H $HOST -p $SENDSMS_PORT`
if [ ` echo $SENDSMS_PORT_RESPONSE | grep 'TCP OK'  | wc -l` -eq 1  ]
then
        ISSENDSMS_PORT_RESPONSE=1  #we are ok
else
        ISSENDSMS_PORT_RESPONSE=0
fi

if [ $ISADMIN_PORT_RESPOND -eq 1 ] && [ $ISSENDSMS_PORT_RESPONSE -eq 1 ]
then
        echo -e "KANNEL: Both ADMIN port $ADMIN_PORT and SENDSMS port $SENDSMS_PORT are alive"
        exit 0;
fi



if [ $ISADMIN_PORT_RESPOND -eq 0 ] && [ $ISSENDSMS_PORT_RESPONSE -eq 0 ]
then
        echo -e "CRITICAL - KANNEL: Both ADMIN port $ADMIN_PORT and SEND SMS port $SENDSMS_PORT ARE NOT ALIVE"
        exit 2;
else
        if [ $ISADMIN_PORT_RESPOND -eq 1 ]
        then
                echo "KANNEL - CRITICAL: SENDSMS port $SENDSMS_PORT is NOT ALIVE"
                exit 1;
        else
                echo " KANNEL - CRITICAL: ADMIN port $ADMIN_PORT is NOT ALIVE"
                exit 1;
        fi
fi