#!/bin/bash
SCRIPT_PATH=/home/monitoring/ASIACELL_RELATED_PROMO/
LOG_PATH=/home/promo/jboss-6.1.0.Final_PLATFORM/server/default/log/
LOG_PATH_ARCHIVE=/home/promo/jboss-6.1.0.Final_PLATFORM/server/default/log/LOGS_ARCHIVE/
PREVIOUS_DAY_PATTERN=`/bin/date -d "$dda 1 days ago" +%g-%m-%d`
FILE=platform-admin.log.$PREVIOUS_DAY_PATTERN
ZIP_FILE=platform-admin.log."$PREVIOUS_DAY_PATTERN".gz
SUBJECT="ASIACELL  BULK MTs"
dda=`date +%Y%m%d`
PREVIOUS_DAY_FORMAT=`/bin/date -d "$dda 1 days ago" +%d/%m/%Y`

NUMBASE_LOG_PATH=/home/promo/jboss-6.1.0.Final_BULK_CONNECTION/server/default/log
NUMBASE_LOG_PATH_ARCHIVE=/home/promo/jboss-6.1.0.Final_BULK_CONNECTION/server/default/log/LOG_ARCHIVE
NUMBASE_FILE=asiacell-numbase-bulk.log.$PREVIOUS_DAY_PATTERN
NUMBASE_ZIP_FILE=asiacell-numbase-bulk.log.$PREVIOUS_DAY_PATTERN.gz


if  [ -f $NUMBASE_LOG_PATH/$NUMBASE_FILE ]
then
	result1=`/bin/grep "NumbaseBulkMTController \[INFO\] Received request" $NUMBASE_LOG_PATH/$NUMBASE_FILE   |awk -F "NumbaseBulkMTController|Body=" {'print $3'} | replace "[" "" | replace "]" "" | replace "}" ""  | tr "," "\n"  |wc -l`
else
	if [ -f $NUMBASE_LOG_PATH_ARCHIVE/$NUMBASE_ZIP_FILE ]
	then
		result1=`/usr/bin/zgrep "NumbaseBulkMTController \[INFO\] Received request"  $NUMBASE_LOG_PATH_ARCHIVE/$NUMBASE_ZIP_FILE  |awk -F "NumbaseBulkMTController|Body=" {'print $3'} | replace "[" "" | replace "]" "" | replace "}" ""  | tr "," "\n"  |wc -l`
	else
		echo "FILE CANNOT BE FOUND" |  /usr/bin/mutt -s  "!!!!!!!!! PROBLEM $SUBJECT !!!!!!!!!!!!!!" n.lazarou@cdialogues.com  p.zafiriadis@cdialogues.com a.floros@cdialogues.com  m.alafouzos@cdialogues.com a.mpalidis@cdialogues.com
	fi
fi

echo $result



##################################### CALCULATION OF PLATFORM BULK MTs  ##############################################################

if  [ -f $LOG_PATH/$FILE ]
then
        result=`/bin/grep TX $LOG_PATH/$FILE | /bin/grep "messageType=BULK_INVITATION" | /bin/sed "s/.*msisdn='\([0-9]*\)'.* bulkId=\([0-9]*\).*/\1 \2/" | /bin/sort -u  | /usr/bin/wc -l`
else
        if [ -f $LOG_PATH_ARCHIVE/$ZIP_FILE ]
       then
          result=`/usr/bin/zgrep TX $LOG_PATH_ARCHIVE/$ZIP_FILE | /bin/grep "messageType=BULK_INVITATION" | /bin/sed "s/.*msisdn='\([0-9]*\)'.* bulkId=\([0-9]*\).*/\1 \2/" | sort -u  | /usr/bin/wc -l`
else
         echo "FILE CANNOT BE FOUND" |  /usr/bin/mutt -s  "!!!!!!!!! PROBLEM $SUBJECT !!!!!!!!!!!!!!" n.lazarou@cdialogues.com  p.zafiriadis@cdialogues.com a.floros@cdialogues.com  m.alafouzos@cdialogues.com a.mpalidis@cdialogues.com
                exit 2;
        fi
fi


##################################### CALCULATION OF NUMBASE BULK MTs  ##############################################################







echo -e "Dear all,
Please be informed that the number of MTs sent yesterday $PREVIOUS_DAY_FORMAT  through Bulk from our PLATFORM (Bulks from our side and Numbase server)  was: $result \n
The Number od MTs sent yesterday $PREVIOUS_DAY_FORMAT  through Bulk from NUMBASE was : $result1 \n
Thank you" |  /usr/bin/mutt -s  "$SUBJECT"  n.lazarou@cdialogues.com  a.sklaveniti@cdialogues.com   l.fourlari@cdialogues.com   k.papadaki@cdialogues.com  g.mitropetrou@cdialogues.com   a.floros@cdialogues.com  p.zafiriadis@cdialogues.com m.alafouzos@cdialogues.com v.kalantzi@cdialogues.com  a.mpalidis@cdialogues.com p.zafiriadis@cdialogues.com






