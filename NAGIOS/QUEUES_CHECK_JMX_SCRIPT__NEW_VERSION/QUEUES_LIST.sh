#!/bin/sh

SCRIPTS_HOME=/home/nagios/monitoring/libexec/CUSTOM/QUEUES_INFORMATION_VERSION2_LIVE_2_USERPROMOFOUR/
##$SCRIPTS_HOME/RETRIEVE_QUEUES_SPAWN.sh > $SCRIPTS_HOME/RETRIEVE_QUEUES_TEMP
HOST=192.168.100.205
PORT=3705
JAVAHOMESCRIPT=/usr/java/jdk1.6.0_26/bin/java







########################## RETRIEVE QUEUES NAMES FROM JMX ###########################################################################
##/bin/cat $SCRIPTS_HOME/RETRIEVE_QUEUES_TEMP  | /bin/grep "type=Queue" | /bin/awk -F "name=" {'print $2'} | /usr/bin/replace  '"' ' ' | /bin/awk {'print $1'} | /bin/grep -v "jms.queue." > $SCRIPTS_HOME/QUEUES_LIST




############################ CREATE SPAWN.sh SCRIPT ######################################################################################## 

echo -e "#!/usr/bin/expect" > $SCRIPTS_HOME/SPAWN.sh
echo -e "\n" >> $SCRIPTS_HOME/SPAWN.sh
echo -e "spawn $JAVAHOMESCRIPT -jar  $SCRIPTS_HOME/jmxterm-1.0-alpha-4-uber.jar  -l $HOST:$PORT"  >> $SCRIPTS_HOME/SPAWN.sh

echo -e "expect \"Welcome to JMX terminal. Type \\\"help\\\" for available commands.\""  >> $SCRIPTS_HOME/SPAWN.sh
echo -e "expect \"$>\"" >> $SCRIPTS_HOME/SPAWN.sh

for i in `cat $SCRIPTS_HOME/QUEUES_LIST`
do
echo -e "send  \"get -b org.hornetq:module=JMS,name=\\\"$i\\\",type=Queue ConsumerCount \\\n\""  >> $SCRIPTS_HOME/SPAWN.sh
echo -e "expect \"$>\""  >> $SCRIPTS_HOME/SPAWN.sh
echo -e "send  \"get -b org.hornetq:module=JMS,name=\\\"$i\\\",type=Queue MessagesAdded \\\n\""  >> $SCRIPTS_HOME/SPAWN.sh
echo -e "expect \"$>\""  >> $SCRIPTS_HOME/SPAWN.sh
echo -e "send  \"get -b org.hornetq:module=JMS,name=\\\"$i\\\",type=Queue MessageCount \\\n\""  >> $SCRIPTS_HOME/SPAWN.sh
echo -e "expect \"$>\""  >> $SCRIPTS_HOME/SPAWN.sh
done

echo -e "send \"exit;\"" >>  $SCRIPTS_HOME/SPAWN.sh


chmod 755 $SCRIPTS_HOME/SPAWN.sh

$SCRIPTS_HOME/SPAWN.sh > $SCRIPTS_HOME/RESULT_SPAWN
wait
/bin/sleep 30;

##/bin/cp  $SCRIPTS_HOME/RESULT_SPAWN_TEMP   $SCRIPTS_HOME/RESULT_SPAWN

