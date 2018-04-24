#!/bin/bash

#####################################
## touch /home/curry/bin/rsync_failover.sh
### % ssh-keygen -t dsa
## ls -ltr ~/.ssh/id_dsa
## cat ~/.ssh/id_dsa.pub | ssh curry@systems.com "cat >> ~/.ssh/authorized_keys"
## chmod 700 authorized_keys  (PRIMARY NODE)
## -rwx------ 1 curry curry 628 Mar  7 13:53 authorized_keys
#####################################

#RSYNC_Failover.sh
## Script Path: /home/curry/bin/rsync_failover.sh
## Script Logs: /home/curry/bin/rsync_failover_[DATE].log
PATH=/usr/local/bin:/bin:/usr/bin
OS_USER=root
RSYNC_PATH=/home/MYSQLOLDDATA
####RSYNC_PATH=/home/kannel_four
PRIMARY_MACHINE=192.168.100.205
LOG_FILE=/root/RSYNC/rsync_`/bin/date +%Y%m%d`.log
echo "##########################################################" >> $LOG_FILE
echo "### Starting Sync Process `/bin/date` ##" >> $LOG_FILE
echo "##########################################################" >> $LOG_FILE


/usr/bin/rsync -av -e ssh "$OS_USER"@"$PRIMARY_MACHINE":"$RSYNC_PATH"/  "$RSYNC_PATH"/  >> "$LOG_FILE" 2>>"$LOG_FILE"

#/usr/bin/rsync   d -av --delete-before -e ssh --exclude  '.vminfo'  --exclude '*VRTSpbx*' --exclude '*lost+found*' "$OS_USER"@"$PRIMARY_MACHINE":"$RSYNC_PATH"/  "$RSYNC_PATH"/  >> "$LOG_FILE" 2>>"$LOG_FILE"
if [ $? -eq 0 ]
        then
                echo -e "\e[32;5m" "### Ending $RSYNC_PATH Sync Process `/bin/date` ###" >> $LOG_FILE
                echo -e -n "\e[0m" >> $LOG_FILE
                echo "########################################" >> $LOG_FILE
                echo >> $LOG_FILE
        else
                echo -e "\e[31;5m" "### ERROR Ending Sync Process `/bin/date` ###" >> $LOG_FILE
                echo -e -n "\e[0m" >> $LOG_FILE
                echo "########################################" >> $LOG_FILE
                echo >> $LOG_FILE

#				for MAILS in vangelis.typaldos@upstreamsystems.com
#                do
#                  Subject_Mail="ERROR - Failover Sync Process of Curry"
#                  Body_Mail="ERROR - Failover Sync Process `/bin/date` of $PRIMARY_MACHINE"
#		           Body_Mail2="`cat $LOG_FILE | grep ERROR -B4 | tail -4`"
#                  echo -e "$Body_Mail\n \n$Body_Mail2" | mail -s "$Subject_Mail" $MAILS
#                done
fi

