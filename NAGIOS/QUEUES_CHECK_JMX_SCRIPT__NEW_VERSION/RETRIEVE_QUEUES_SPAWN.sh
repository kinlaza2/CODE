#!/usr/bin/expect

spawn /usr/java/jdk1.6.0_26/bin/java -jar  /home/nagios/monitoring/libexec/CUSTOM/QUEUES_INFORMATION_VERSION2_LIVE_2_USERPROMOFOUR/jmxterm-1.0-alpha-4-uber.jar -l 192.168.100.205:3705
expect "Welcome to JMX terminal. Type \"help\" for available commands."
expect "$>"
send  "beans \n"
expect "$>"
send "exit;"
