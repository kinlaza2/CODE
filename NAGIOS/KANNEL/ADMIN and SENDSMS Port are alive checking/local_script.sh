#! /bin/bash
USER=nagios
PASSWORD=nagios
LIBEXEC=/home/nagios/nagios/libexec
REMOTE_EXEC=/home/nagios/monitoring/libexec/CUSTOM

$LIBEXEC/check_by_ssh -H  $1  -t 70 -q -C "$REMOTE_EXEC/check_kannel_ports $2 $3 $4"
