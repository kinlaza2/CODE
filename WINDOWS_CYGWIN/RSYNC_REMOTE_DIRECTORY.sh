## OPEN SSH KEYS so can access through SSH with no password

LOCAL_DIRECTORY=/cygdrive/d/CD/BACKUP_CD_INFRA/60/NAGIOS/
REMOTE_SERVER=10.1.1.60
REMOTE_DIRECTORY=/home/nagios/TEMP/
REMOTE_USER=root
SCRIPT_PATH=/home/ni_new/SYNC_SERVERS_JOB/
LOG_FILE=NAGIOS_SYN.log


echo -e "-------------------------- PROCESS START ON `date` -------------------------------------------------------------------------------------------" >> $SCRIPT_PATH/$LOG_FILE
rsync -azPv --delete $REMOTE_USER@$REMOTE_SERVER:$REMOTE_DIRECTORY   $LOCAL_DIRECTORY >> $SCRIPT_PATH/$LOG_FILE
echo -e "-------------------------- PREOCESS FINISHED ON `date` ---------------------------------------------------------------------------------------" >> $SCRIPT_PATH/$LOG_FILE


