#!/bin/bash

USERNAME=administrator
PASSWORD=administrator123enter
CAS_PAGE=https://lebanonpromo.cdialogues.com/cas/login
SCRIPT_PATH=/home/lnikos/SCRIPTS
PROMO_NAME=LEBANON
STATE_FILE=DB
#WHERE DB FIELDS ARE : BULK_ID,STATUS,BULK_NAME,UPLOAD_ID,LAST_NOTIFICATION_TIME,NUMBER_OF_NOTIFICATIONS


###############################################  FUNCTIONS   ##################################################################
function additem()
{

{

function removeitem()
{

{

function updateitem()
{

{

function bulkexistonlist()
{

{


###############################################################################################################################

################################ WGET THE BULK PAGE and save exit on TEMP file###############################
cd $SCRIPT_PATH
#wget -a log.txt -O loginPage.html -S --keep-session-cookies --save-cookies cookies1.txt $CAS_PAGE
#LT=`grep 'name="lt"' loginPage.html | sed 's/.*value="\(.*\)".*/\1/'`
#EXECUTION=`grep 'name="execution"' loginPage.html | sed 's/.*value="\(.*\)".*/\1/'`
#EVENT_ID=submit
#POSTDATA=username=$USERNAME'&'password=$PASSWORD'&'lt=$LT'&'execution=${EXECUTION}'&'_eventId=$EVENT_ID
#wget -a log.txt -O -  --post-data $POSTDATA --keep-session-cookies --load-cookies cookies1.txt --save-cookies cookies2.txt $CAS_PAGE  > /dev/null
#wget -a log.txt -O - --keep-session-cookies --load-cookies cookies2.txt --save-cookies cookies3.txt https://lebanonpromo.cdialogues.com/platform-webapp/admin/bulk/list?campaignId=1> $SCRIPT_PATH/TEMPFILE


######################  ANALYZE BULK PAGE #############################################################################
############ FIND IF BULK IS CONFIGURED TODAY AND IF IT RUNNING #########3
##TODAY=`date +%d/%m/%Y`
TODAY="16/04/2016"                                                               `
echo $TODAY



grep   $TODAY $SCRIPT_PATH/TEMPFILE   > $SCRIPT_PATH/TEMPFILE2


cat $SCRIPT_PATH/TEMPFILE2  | awk -F "bulk-" {'print $2'} |  awk -F "-" {'print $1'} > $SCRIPT_PATH/BULKIDS

BULK_IDS_LINES=`grep [0-9]  $SCRIPT_PATH/BULKIDS | wc -l`
echo $BULK_IDS_LINES

## !! PENDING isnumeric   BULK_IDS_LINES

if [ $BULK_IDS_LINES -eq 0 ]
then
    echo "NO BULK FOR TODAY HAS BEEN SETUP"
else
    for i in `cat $SCRIPT_PATH/BULKIDS`
    do
        BULK_ID=$i
        echo $BULK_ID
        BULKNAME=`grep bulk-$i $SCRIPT_PATH/TEMPFILE   | grep name | replace ">" "###" | replace  "<" "###" | awk -F "###" {'print $3'}`
        echo $BULKNAME
        STATUS=`grep -A 4  bulk-$i-status $SCRIPT_PATH/TEMPFILE  | grep span | replace "span" "" | replace "<" "" | replace ">" "" | replace "/" "" | awk {'print $1'}`
        echo $STATUS
        ADDRESSED=`grep bulk-$i-addressed $SCRIPT_PATH/TEMPFILE  |  awk -F ">" {'print $2'} | awk -F "<" {'print $1'}`
        echo $ADDRESSED
        RESPONSE_RATE=`grep bulk-$i-responseRate $SCRIPT_PATH/TEMPFILE | awk -F ">" {'print $2'} | awk -F "<" {'print $1'}`
        echo $RESPONSE_RATE
        UPLOAD_ID=`grep -A 5 "bulk-$i-uploadIds" $SCRIPT_PATH/TEMPFILE | grep div | awk -F "div>" {'print $2'} | replace "</" "" | awk {'print $1'}`
        echo $UPLOAD_ID

        echo $STATUS

        case $STATUS in
        RUNNING)
                  echo "Bulk with ID $BULK_ID is running"
                  ;;
        STOPPED)
                  echo STOPPED
                        ;;
        DRAFT)
                  echo -e "$PROMO_NAME - INFO - BULK $BULKNAME HAS BEEN FOUND  IN DRAFT STATE FOR $DATE"
                        ;;
        CONFIGURED)
                  echo -e "$PROMO_NAME - ATTENTION - BULK $BULKNAME HAS BEEN FOUND TO BE IN CONFIGURED STATE FOR $DATE !!"
                        ;;
        TESTING)

                        ;;
        *)
                echo  -e "TECHNICAL ISSUE ON SCRIPT FOR BULK RUN DETECTION ON $PROMO_NAME"
                exit 2;
                ;;
        esac


    done
fi

#rm -f    $SCRIPT_PATH/TEMPFILE
#rm -f    $SCRIPT_PATH/TEMPFILE2
#rm -f    $SCRIPT_PATH/BULKIDS
