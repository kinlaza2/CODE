#!/bin/bash

SCRIPT_PATH=/home/kannel/CLEAR_DB
USER_PATH=/home/kannel
TWO_DAYS_AGO_PATTERN=`/bin/date -d "$dda 2 days ago" +%g%m%d`
dda=`date +%Y%m%d`
LOG_FILE=$SCRIPT_PATH/LOG_FILE.log
MYSQLPATH="/home/kannel//MYSQL_5_5_23/bin/mysql  --socket=/home/kannel/MYSQL_5_5_23/data/mysql.sock --default-character-set=utf8"

/bin/echo $TWO_DAYS_AGO_PATTERN

/bin/echo -e "`date` || ##########################################################################################" >> $LOG_FILE
/bin/echo -e "`date` || ######			PROCESS START FOR `date` 				  #####" >> $LOG_FILE


/bin/echo -e "`date` || PROCESS START" >> $LOG_FILE
if [ -f $SCRIPT_PATH/TEMPO.sql ]
then
	/bin/rm $SCRIPT_PATH/TEMPO.sql
fi

/bin/echo -e "`date` || RECORDS FOR $TWO_DAYS_AGO_PATTERN WILL BE DELETED" >> $LOG_FILE

/bin/echo -e " select *   from dlr where ts like '____________AAAAA____________';  " | /home/kannel/MYSQL_5_5_23/bin/replace AAAAA  $TWO_DAYS_AGO_PATTERN > $SCRIPT_PATH/TEMPO.sql
$MYSQLPATH  -u root -pCDialogues123enter -D KANNEL_DB < $SCRIPT_PATH/TEMPO.sql >> $LOG_FILE

/bin/echo -e "`date` || ######                       PROCESS END FOR `date`                                  #####" >> $LOG_FILE
/bin/echo -e "`date` || ##########################################################################################" >> $LOG_FILE
/bin/echo -e "\n" >> $LOG_FILE
